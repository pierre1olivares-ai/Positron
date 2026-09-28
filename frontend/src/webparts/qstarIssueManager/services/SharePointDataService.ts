import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders/web";
import "@pnp/sp/folders/list";
import type { IItems } from "@pnp/sp/items";

import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, DEFAULT_SETTINGS, normalizeSettings } from "../models/ISettings";
import {
  ISSUE_FIELDS,
  PROGRESS_FIELDS,
  ISSUE_PERSON_FIELD_NAMES,
  ISSUE_SELECT_FIELD_NAMES,
  PROGRESS_SELECT_FIELD_NAMES,
  DEFAULT_ISSUES_LIST,
  DEFAULT_PROGRESS_LIST,
} from "./fieldMap";
import { readPersonValue } from "./sharePointValues";
import { normalizeDateOnly } from "../domain/calendarDates";
import { normalizeRegion } from "../domain/referenceData";
import { IssueConflictError, IssueRefreshError } from "./issueErrors";

type SPItem = Record<string, unknown> & { Id: number };

const CONFIG_LIST = "Q-Star Config";
const CONFIG_JSON_FIELD = "SettingsJson";
const REFERENCE_OFFSET_FIELD = "ReferenceOffset";

const yesNoToBool = (v: string | undefined): boolean => v === "Yes";
const boolToYesNo = (v: boolean | undefined): string => (v ? "Yes" : "No");
// SharePoint REST requires null (rather than an empty string) to clear DateTime fields.
// eslint-disable-next-line @rushstack/no-new-null
const emptyToNull = (v: string): string | null => v || null;

/**
 * Real production data layer: reads/writes the "Q-Star Issues" and
 * "Q-Star Progress Log" SharePoint lists via SharePoint REST (PnPjs),
 * using the signed-in user's session — no Entra app registration or
 * Graph admin consent required, since the web part is hosted on the
 * same site as the lists. Pass `siteUrl` only if the lists live on a
 * different site than where the web part runs.
 */
export class SharePointDataService implements IDataService {
  private sp: SPFI;
  private referenceOffset?: number;
  private progressRoot?: string;
  private webOrigin?: string;
  private readonly snapshots = new Map<number, IIssue>();

  constructor(
    context: WebPartContext,
    private issuesListName: string = DEFAULT_ISSUES_LIST,
    private progressListName: string = DEFAULT_PROGRESS_LIST,
    siteUrl?: string,
    client?: SPFI
  ) {
    this.sp = client || (siteUrl
      ? spfi(siteUrl).using(SPFx(context))
      : spfi().using(SPFx(context)));
  }

  public async loadIssues(): Promise<IIssue[]> {
    await Promise.all([this.loadReferenceOffset(), this.getProgressRoot()]);
    const [items, progressItems]: [SPItem[], SPItem[]] = await Promise.all([
      this.loadAll(this.sp.web.lists
        .getByTitle(this.issuesListName)
        .items.select("Id", ...ISSUE_SELECT_FIELD_NAMES)
        .expand(...ISSUE_PERSON_FIELD_NAMES)),
      this.loadAll(this.sp.web.lists
        .getByTitle(this.progressListName)
        .items.select("Id", ...PROGRESS_SELECT_FIELD_NAMES)
        .expand(PROGRESS_FIELDS.author).filter("FSObjType eq 0")),
    ]);

    const logsByParent = new Map<number, IProgressLogEntry[]>();
    for (const p of progressItems) {
      // Folder ACLs are the security boundary; a writable ParentItemId is not.
      const parentId = this.progressParentId(p);
      if (parentId === undefined) continue;
      const entry = this.toProgressEntry(p);
      const list = logsByParent.get(parentId) || [];
      list.push(entry);
      logsByParent.set(parentId, list);
    }
    logsByParent.forEach((entries) => entries.sort((a, b) => a.ts.localeCompare(b.ts)));

    return items.map((item) => this.remember(this.toIssue(item, logsByParent.get(item.Id) || [])));
  }

  public async getIssue(id: number): Promise<IIssue> {
    await this.loadReferenceOffset();
    const root = await this.getProgressRoot();
    const [item, entries] = await Promise.all([
      this.sp.web.lists.getByTitle(this.issuesListName).items.getById(id)
        .select("Id", ...ISSUE_SELECT_FIELD_NAMES).expand(...ISSUE_PERSON_FIELD_NAMES)(),
      this.loadAll(this.sp.web.lists.getByTitle(this.progressListName).items
        .select("Id", ...PROGRESS_SELECT_FIELD_NAMES).expand(PROGRESS_FIELDS.author)
        .filter(`FSObjType eq 0 and FileDirRef eq '${escapeOData(`${root}/issue-${id}`)}'`)),
    ]);
    const progress = entries.filter((entry) => this.progressParentId(entry) === id)
      .map((entry) => this.toProgressEntry(entry)).sort((a, b) => a.ts.localeCompare(b.ts));
    return this.remember(this.toIssue(item as SPItem, progress));
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    const offset = await this.loadReferenceOffset();
    const fields = await this.toSPFields({ ...issue, qsNumber: undefined });
    const result = await this.sp.web.lists
      .getByTitle(this.issuesListName)
      .items.add(fields);
    const id = result.Id as number;
    const qsNumber = offset + id;
    // Once POST succeeds, never turn a follow-up failure into a retryable create.
    let saveWarning: string | undefined;
    try {
      await this.sp.web.lists.getByTitle(this.issuesListName).items.getById(id)
        .update({ [ISSUE_FIELDS.qsNumber]: qsNumber });
    } catch {
      saveWarning = "The report was saved. Its reference is reserved, but SharePoint reference synchronization is pending.";
    }
    try {
      return { ...await this.getIssue(id), saveWarning };
    } catch {
      const fallback = this.toIssue({
        ...fields, ...result, Id: id, [ISSUE_FIELDS.qsNumber]: qsNumber,
        [ISSUE_FIELDS.createdBy]: { Id: fields.ReportedById, Title: issue.createdBy, EMail: issue.createdByEmail },
        [ISSUE_FIELDS.taskOwner]: { Id: fields.TaskOwnerId, Title: issue.taskOwner, EMail: issue.taskOwnerEmail },
        [ISSUE_FIELDS.verifiedBy]: { Id: fields.VerifiedById, Title: issue.verifiedBy, EMail: issue.verifiedByEmail },
      } as SPItem, []);
      return { ...fallback, eTag: undefined, saveWarning: "The report was saved. Refresh to load its latest details; do not submit it again." };
    }
  }

  public async updateIssue(id: number, patch: Partial<IIssue>, expectedETag?: string): Promise<IIssue> {
    const previous = this.snapshots.get(id);
    const eTag = expectedETag || (previous && previous.eTag);
    if (!eTag || eTag === "*") throw new Error("Refresh this issue before saving so its current version can be checked.");
    const changed: Partial<IIssue> = {};
    Object.keys(patch).forEach((key) => {
      const name = key as keyof IIssue;
      if (!previous || previous.eTag !== eTag || patch[name] !== previous[name]) {
        (changed as Record<string, unknown>)[name] = patch[name];
      }
    });
    // References are immutable after allocation; callers cannot renumber issues.
    delete changed.qsNumber;
    const fields = await this.toSPFields(changed);
    if (Object.keys(fields).length === 0) return this.getIssue(id);
    try {
      await this.sp.web.lists.getByTitle(this.issuesListName).items.getById(id).update(fields, eTag);
    } catch (error) {
      if (httpStatus(error) !== 412) throw error;
      let freshIssue: IIssue | undefined;
      try { freshIssue = await this.getIssue(id); } catch { /* Keep the conflict even if the refresh is unavailable. */ }
      throw new IssueConflictError(id, freshIssue);
    }
    this.snapshots.delete(id);
    try { return await this.getIssue(id); }
    catch { throw new IssueRefreshError(id); }
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<IProgressLogEntry> {
    if (!entry.text.trim()) throw new Error("Enter a progress update before posting.");
    const root = await this.getProgressRoot();
    const folder = `${root}/issue-${id}`;
    const list = this.sp.web.lists.getByTitle(this.progressListName);
    try {
      await this.sp.web.getFolderByServerRelativePath(folder).select("Exists")();
    } catch (error) {
      if (httpStatus(error) !== 404) throw error;
      try { await list.rootFolder.folders.addUsingPath(folder); }
      catch {
        throw new Error("Progress access for this issue is still being prepared. Your text has been kept; retry after the assignment permissions flow completes.");
      }
    }
    const result = await list.addValidateUpdateItemUsingPath([
      { FieldName: "Title", FieldValue: `Issue ${id} progress` },
      { FieldName: PROGRESS_FIELDS.text, FieldValue: entry.text },
    ], `${this.webOrigin}${folder}`);
    const failures = result.filter((field) => field.HasException);
    if (failures.length) throw new Error(failures.map((field) => field.ErrorMessage || "Progress entry validation failed.").join(" "));
    const idResult = result.filter((field) => field.ItemId || field.FieldName === "Id" || field.FieldName === "ID")[0];
    const entryId = idResult && (idResult.ItemId || Number(idResult.FieldValue));
    if (entryId) {
      try {
        const saved = await list.items.getById(entryId).select("Id", ...PROGRESS_SELECT_FIELD_NAMES).expand(PROGRESS_FIELDS.author)();
        return this.toProgressEntry(saved as SPItem);
      } catch { /* The append succeeded; a failed read must not encourage a duplicate append. */ }
    }
    return { ...entry, id: entryId || undefined, saveWarning: "Your update was posted. Refresh to reload its server-recorded author and time; do not post it again." };
  }

  public async loadSettings(): Promise<ISettings> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(CONFIG_LIST)
        .items.select(CONFIG_JSON_FIELD)
        .top(1)();
      if (items.length && items[0][CONFIG_JSON_FIELD]) {
        const stored = items[0][CONFIG_JSON_FIELD] as string;
        return normalizeSettings(JSON.parse(stored));
      }
    } catch (error) {
      const message = errorMessage(error);
      if (/404|not found|does not exist/i.test(message)) return { ...DEFAULT_SETTINGS };
      throw new Error(`Could not load Q-Star settings: ${message}`);
    }
    return { ...DEFAULT_SETTINGS };
  }

  public async saveSettings(settings: ISettings): Promise<void> {
    const list = this.sp.web.lists.getByTitle(CONFIG_LIST);
    const items = await list.items.select("Id").top(1)();
    const json = JSON.stringify(normalizeSettings(settings));
    if (items.length) {
      await list.items.getById(items[0].Id).update({ [CONFIG_JSON_FIELD]: json });
    } else {
      await list.items.add({ Title: "Q-Star Settings", [CONFIG_JSON_FIELD]: json });
    }
  }

  private async loadReferenceOffset(): Promise<number> {
    if (this.referenceOffset !== undefined) return this.referenceOffset;
    const items = await this.sp.web.lists.getByTitle(CONFIG_LIST).items.select(REFERENCE_OFFSET_FIELD).top(1)();
    const offset = items.length ? Number(items[0][REFERENCE_OFFSET_FIELD]) : NaN;
    if (!isFinite(offset) || Math.floor(offset) !== offset || offset < 1000 || offset > 9007199254740991) {
      throw new Error("Q-Star reference allocation is not configured. Run the current provisioning script before creating or loading issues.");
    }
    this.referenceOffset = offset;
    return offset;
  }

  private async getProgressRoot(): Promise<string> {
    if (this.progressRoot) return this.progressRoot;
    const [info, web] = await Promise.all([
      this.sp.web.lists.getByTitle(this.progressListName)
        .select("RootFolder/ServerRelativeUrl").expand("RootFolder")(),
      this.sp.web.select("Url")(),
    ]);
    this.progressRoot = (info.RootFolder.ServerRelativeUrl as string).replace(/\/$/, "");
    this.webOrigin = new URL(web.Url as string).origin;
    return this.progressRoot;
  }

  private progressParentId(item: SPItem): number | undefined {
    const directory = item.FileDirRef as string;
    if (!directory || !this.progressRoot || directory.indexOf(`${this.progressRoot}/issue-`) !== 0) return undefined;
    const suffix = directory.substring(this.progressRoot.length + "/issue-".length);
    return /^[1-9]\d*$/.test(suffix) ? Number(suffix) : undefined;
  }

  private toProgressEntry(item: SPItem): IProgressLogEntry {
    const author = readPersonValue(item[PROGRESS_FIELDS.author] as never);
    return {
      id: item.Id,
      ts: (item.Created as string) || (item[PROGRESS_FIELDS.entryDate] as string) || "",
      author: author.displayName,
      authorId: author.id,
      authorEmail: author.email,
      text: (item[PROGRESS_FIELDS.text] as string) || "",
    };
  }

  private remember(issue: IIssue): IIssue {
    this.snapshots.set(issue.id, issue);
    return issue;
  }

  private toIssue(item: SPItem, progressLog: IProgressLogEntry[]): IIssue {
    const f = ISSUE_FIELDS;
    const createdBy = readPersonValue(item[f.createdBy] as never);
    const taskOwner = readPersonValue(item[f.taskOwner] as never);
    const verifiedBy = readPersonValue(item[f.verifiedBy] as never);
    return {
      id: item.Id as number,
      eTag: (item["odata.etag"] || item["@odata.etag"] || (item.__metadata as { etag?: string } | undefined)?.etag) as string | undefined,
      qsNumber: typeof item[f.qsNumber] === "number" ? item[f.qsNumber] as number : (this.referenceOffset as number) + item.Id,
      triaged: yesNoToBool(item[f.triaged] as string),
      status: (item[f.status] as IIssue["status"]) || undefined,
      taskCreated: (item[f.taskCreated] as IIssue["taskCreated"]) || "No",
      transformedInto: (item[f.transformedInto] as IIssue["transformedInto"]) || undefined,

      shortSummary: (item[f.shortSummary] as string) || "",
      description: (item[f.description] as string) || "",
      immediateAction: (item[f.immediateAction] as string) || "",
      severity: (item[f.severity] as IIssue["severity"]) || "Medium",
      createdBy: createdBy.displayName,
      createdById: createdBy.id,
      createdByEmail: createdBy.email,
      reportDate: normalizeDateOnly(item[f.reportDate] as string),
      departmentBU: (item[f.departmentBU] as string) || "",
      region: normalizeRegion((item[f.region] as string) || ""),
      alreadyInContact: (item[f.alreadyInContact] as IIssue["alreadyInContact"]) || "No",
      deviationType: (item[f.deviationType] as string) || "",
      issueOrigin: (item[f.issueOrigin] as string) || "",
      additionalComments: (item[f.additionalComments] as string) || "",

      followUp: (item[f.followUp] as string) || "",
      taskOwner: taskOwner.displayName,
      taskOwnerId: taskOwner.id,
      taskOwnerEmail: taskOwner.email,
      ownerBU: (item[f.ownerBU] as string) || (item[f.departmentBU] as string) || "",
      dueDate: normalizeDateOnly(item[f.dueDate] as string),

      rootCause: (item[f.rootCause] as string) || "",
      correctiveAction: (item[f.correctiveAction] as string) || "",
      implementationDate: normalizeDateOnly(item[f.implementationDate] as string),
      effectivenessCheck: (item[f.effectivenessCheck] as string) || "",
      verifiedBy: verifiedBy.displayName,
      verifiedById: verifiedBy.id,
      verifiedByEmail: verifiedBy.email,
      verifiedDate: normalizeDateOnly(item[f.verifiedDate] as string),
      closedDate: normalizeDateOnly(item[f.closedDate] as string),
      closedAt: (item[f.closedAt] as string) || "",

      holdReason: (item[f.holdReason] as string) || "",
      holdUntil: normalizeDateOnly(item[f.holdUntil] as string),

      ownerUpdate: yesNoToBool(item[f.ownerUpdate] as string),
      ownerUpdateAt: (item[f.ownerUpdateAt] as string) || "",
      ownerUpdateText: (item[f.ownerUpdateText] as string) || "",
      reminderCycle: (item[f.reminderCycle] as string) || "",

      attachments: [],
      progressLog,
    };
  }

  /** Converts only the keys present on `patch`/`issue` into SharePoint internal field names. */
  private async toSPFields(issue: Partial<IIssue>): Promise<Record<string, unknown>> {
    const f = ISSUE_FIELDS;
    const out: Record<string, unknown> = {};
    const set = <K extends keyof IIssue>(key: K, spName: string, transform?: (v: IIssue[K]) => unknown): void => {
      if (issue[key] === undefined) return;
      out[spName] = transform ? transform(issue[key] as IIssue[K]) : issue[key];
    };

    set("qsNumber", f.qsNumber);
    set("shortSummary", f.shortSummary);
    if (issue.shortSummary !== undefined) out.Title = issue.shortSummary || "Q-Star Issue";
    set("description", f.description);
    set("immediateAction", f.immediateAction);
    set("severity", f.severity);
    set("reportDate", f.reportDate, emptyToNull);
    set("departmentBU", f.departmentBU);
    set("region", f.region);
    set("alreadyInContact", f.alreadyInContact);
    set("deviationType", f.deviationType);
    set("issueOrigin", f.issueOrigin);
    set("additionalComments", f.additionalComments);

    set("followUp", f.followUp);
    set("status", f.status);
    set("transformedInto", f.transformedInto);
    set("taskCreated", f.taskCreated);
    set("triaged", f.triaged, boolToYesNo);

    set("ownerBU", f.ownerBU);
    set("dueDate", f.dueDate, emptyToNull);

    set("rootCause", f.rootCause);
    set("correctiveAction", f.correctiveAction);
    set("implementationDate", f.implementationDate, emptyToNull);
    set("effectivenessCheck", f.effectivenessCheck);
    set("verifiedDate", f.verifiedDate, emptyToNull);
    set("closedDate", f.closedDate, emptyToNull);
    set("closedAt", f.closedAt, emptyToNull);

    set("holdReason", f.holdReason);
    set("holdUntil", f.holdUntil, emptyToNull);

    set("ownerUpdate", f.ownerUpdate, boolToYesNo);
    set("ownerUpdateAt", f.ownerUpdateAt, emptyToNull);
    set("ownerUpdateText", f.ownerUpdateText);
    set("reminderCycle", f.reminderCycle);

    await Promise.all([
      this.setPersonField(out, f.createdBy, issue.createdById, issue.createdByEmail, issue.createdBy),
      this.setPersonField(out, f.taskOwner, issue.taskOwnerId, issue.taskOwnerEmail, issue.taskOwner),
      this.setPersonField(out, f.verifiedBy, issue.verifiedById, issue.verifiedByEmail, issue.verifiedBy),
    ]);

    return out;
  }

  private async setPersonField(
    out: Record<string, unknown>,
    fieldName: string,
    id: number | undefined,
    email: string | undefined,
    displayName: string | undefined
  ): Promise<void> {
    const idField = `${fieldName}Id`;
    // An edited email must resolve even if an old cached lookup ID accompanies it.
    if (email) {
      const ensured = await this.sp.web.ensureUser(email);
      out[idField] = ensured.Id;
      return;
    }
    if (id !== undefined) {
      out[idField] = id || null;
      return;
    }
    if (displayName === "") {
      out[idField] = null;
    }
  }

  private async loadAll(query: IItems): Promise<SPItem[]> {
    const all: SPItem[] = [];
    for await (const page of query.top(2000)) {
      all.push(...(page as SPItem[]));
    }
    return all;
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function httpStatus(error: unknown): number | undefined {
  const candidate = error as { status?: number; response?: { status?: number } };
  return candidate && (candidate.status || (candidate.response && candidate.response.status));
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}
