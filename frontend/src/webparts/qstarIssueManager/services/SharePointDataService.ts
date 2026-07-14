import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import type { IItems } from "@pnp/sp/items";

import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, DEFAULT_SETTINGS } from "../models/ISettings";
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

type SPItem = Record<string, unknown> & { Id: number };

const CONFIG_LIST = "Q-Star Config";
const CONFIG_JSON_FIELD = "SettingsJson";

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

  constructor(
    context: WebPartContext,
    private issuesListName: string = DEFAULT_ISSUES_LIST,
    private progressListName: string = DEFAULT_PROGRESS_LIST,
    siteUrl?: string
  ) {
    this.sp = siteUrl
      ? spfi(siteUrl).using(SPFx(context))
      : spfi().using(SPFx(context));
  }

  public async loadIssues(): Promise<IIssue[]> {
    const [items, progressItems]: [SPItem[], SPItem[]] = await Promise.all([
      this.loadAll(this.sp.web.lists
        .getByTitle(this.issuesListName)
        .items.select("Id", ...ISSUE_SELECT_FIELD_NAMES)
        .expand(...ISSUE_PERSON_FIELD_NAMES)),
      this.loadAll(this.sp.web.lists
        .getByTitle(this.progressListName)
        .items.select("Id", ...PROGRESS_SELECT_FIELD_NAMES)
        .expand(PROGRESS_FIELDS.author)),
    ]);

    const logsByParent = new Map<number, IProgressLogEntry[]>();
    for (const p of progressItems) {
      const parentId = p[PROGRESS_FIELDS.parentItemId] as number;
      const author = readPersonValue(p[PROGRESS_FIELDS.author] as never);
      const entry: IProgressLogEntry = {
        ts: p[PROGRESS_FIELDS.entryDate] as string,
        author: author.displayName,
        authorId: author.id,
        authorEmail: author.email,
        text: (p[PROGRESS_FIELDS.text] as string) || "",
      };
      const list = logsByParent.get(parentId) || [];
      list.push(entry);
      logsByParent.set(parentId, list);
    }
    logsByParent.forEach((entries) => entries.sort((a, b) => a.ts.localeCompare(b.ts)));

    return items.map((item) => this.toIssue(item, logsByParent.get(item.Id) || []));
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    const fields = await this.toSPFields(issue);
    const result = await this.sp.web.lists
      .getByTitle(this.issuesListName)
      .items.add(fields);
    const created = await this.sp.web.lists
      .getByTitle(this.issuesListName)
      .items.getById(result.Id as number)
      .select("Id", ...ISSUE_SELECT_FIELD_NAMES)
      .expand(...ISSUE_PERSON_FIELD_NAMES)();
    return this.toIssue(created as SPItem, []);
  }

  public async updateIssue(id: number, patch: Partial<IIssue>): Promise<void> {
    const fields = await this.toSPFields(patch);
    if (Object.keys(fields).length === 0) return;
    await this.sp.web.lists
      .getByTitle(this.issuesListName)
      .items.getById(id)
      .update(fields);
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<void> {
    const authorId = entry.authorId || (await this.sp.web.currentUser.select("Id")()).Id;
    await this.sp.web.lists.getByTitle(this.progressListName).items.add({
      Title: `Issue ${id} progress`,
      [PROGRESS_FIELDS.parentItemId]: id,
      [`${PROGRESS_FIELDS.author}Id`]: authorId,
      [PROGRESS_FIELDS.entryDate]: entry.ts,
      [PROGRESS_FIELDS.text]: entry.text,
    });
  }

  public async loadSettings(): Promise<ISettings> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(CONFIG_LIST)
        .items.select(CONFIG_JSON_FIELD)
        .top(1)();
      if (items.length && items[0][CONFIG_JSON_FIELD]) {
        const stored = items[0][CONFIG_JSON_FIELD] as string;
        return { ...DEFAULT_SETTINGS, ...JSON.parse(stored) } as ISettings;
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
    const json = JSON.stringify(settings);
    if (items.length) {
      await list.items.getById(items[0].Id).update({ [CONFIG_JSON_FIELD]: json });
    } else {
      await list.items.add({ Title: "Q-Star Settings", [CONFIG_JSON_FIELD]: json });
    }
  }

  private toIssue(item: SPItem, progressLog: IProgressLogEntry[]): IIssue {
    const f = ISSUE_FIELDS;
    const createdBy = readPersonValue(item[f.createdBy] as never);
    const taskOwner = readPersonValue(item[f.taskOwner] as never);
    const verifiedBy = readPersonValue(item[f.verifiedBy] as never);
    return {
      id: item.Id as number,
      qsNumber: item[f.qsNumber] as number,
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
      reportDate: (item[f.reportDate] as string) || "",
      departmentBU: (item[f.departmentBU] as string) || "",
      region: (item[f.region] as string) || "",
      alreadyInContact: (item[f.alreadyInContact] as IIssue["alreadyInContact"]) || "No",
      deviationType: (item[f.deviationType] as string) || "",
      issueOrigin: (item[f.issueOrigin] as string) || "",
      additionalComments: (item[f.additionalComments] as string) || "",

      followUp: (item[f.followUp] as string) || "",
      taskOwner: taskOwner.displayName,
      taskOwnerId: taskOwner.id,
      taskOwnerEmail: taskOwner.email,
      ownerBU: (item[f.ownerBU] as string) || (item[f.departmentBU] as string) || "",
      dueDate: (item[f.dueDate] as string) || "",

      rootCause: (item[f.rootCause] as string) || "",
      correctiveAction: (item[f.correctiveAction] as string) || "",
      implementationDate: (item[f.implementationDate] as string) || "",
      effectivenessCheck: (item[f.effectivenessCheck] as string) || "",
      verifiedBy: verifiedBy.displayName,
      verifiedById: verifiedBy.id,
      verifiedByEmail: verifiedBy.email,
      verifiedDate: (item[f.verifiedDate] as string) || "",
      closedDate: (item[f.closedDate] as string) || "",
      closedAt: (item[f.closedAt] as string) || "",

      holdReason: (item[f.holdReason] as string) || "",
      holdUntil: (item[f.holdUntil] as string) || "",

      ownerUpdate: yesNoToBool(item[f.ownerUpdate] as string),
      ownerUpdateAt: (item[f.ownerUpdateAt] as string) || "",
      ownerUpdateText: (item[f.ownerUpdateText] as string) || "",

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
    if (id !== undefined) {
      out[idField] = id || null;
      return;
    }
    if (email) {
      const ensured = await this.sp.web.ensureUser(email);
      out[idField] = ensured.Id;
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
