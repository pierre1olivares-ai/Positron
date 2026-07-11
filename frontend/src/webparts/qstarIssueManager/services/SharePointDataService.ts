import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, DEFAULT_SETTINGS } from "../models/ISettings";
import {
  ISSUE_FIELDS,
  PROGRESS_FIELDS,
  ISSUE_FIELD_NAMES,
  PROGRESS_FIELD_NAMES,
  DEFAULT_ISSUES_LIST,
  DEFAULT_PROGRESS_LIST,
} from "./fieldMap";

type SPItem = Record<string, unknown> & { Id: number };

const CONFIG_LIST = "Q-Star Config";
const CONFIG_JSON_FIELD = "SettingsJson";

const yesNoToBool = (v: string | undefined): boolean => v === "Yes";
const boolToYesNo = (v: boolean | undefined): string => (v ? "Yes" : "No");

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
      this.sp.web.lists
        .getByTitle(this.issuesListName)
        .items.select("Id", ...ISSUE_FIELD_NAMES)
        .top(5000)(),
      this.sp.web.lists
        .getByTitle(this.progressListName)
        .items.select("Id", ...PROGRESS_FIELD_NAMES)
        .top(5000)(),
    ]);

    const logsByParent = new Map<number, IProgressLogEntry[]>();
    for (const p of progressItems) {
      const parentId = p[PROGRESS_FIELDS.parentItemId] as number;
      const entry: IProgressLogEntry = {
        ts: p[PROGRESS_FIELDS.entryDate] as string,
        author: (p[PROGRESS_FIELDS.author] as string) || "",
        text: (p[PROGRESS_FIELDS.text] as string) || "",
      };
      const list = logsByParent.get(parentId) || [];
      list.push(entry);
      logsByParent.set(parentId, list);
    }

    return items.map((item) => this.toIssue(item, logsByParent.get(item.Id) || []));
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    const fields = this.toSPFields(issue);
    const result = await this.sp.web.lists
      .getByTitle(this.issuesListName)
      .items.add(fields);
    return this.toIssue(result as unknown as SPItem, []);
  }

  public async updateIssue(id: number, patch: Partial<IIssue>): Promise<void> {
    const fields = this.toSPFields(patch);
    if (Object.keys(fields).length === 0) return;
    await this.sp.web.lists
      .getByTitle(this.issuesListName)
      .items.getById(id)
      .update(fields);
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<void> {
    await this.sp.web.lists.getByTitle(this.progressListName).items.add({
      [PROGRESS_FIELDS.parentItemId]: id,
      [PROGRESS_FIELDS.author]: entry.author,
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
        return { ...DEFAULT_SETTINGS, ...JSON.parse(items[0][CONFIG_JSON_FIELD]) };
      }
    } catch {
      // Config list not provisioned yet — fall back to defaults.
    }
    return DEFAULT_SETTINGS;
  }

  public async saveSettings(settings: ISettings): Promise<void> {
    const list = this.sp.web.lists.getByTitle(CONFIG_LIST);
    const items = await list.items.select("Id").top(1)();
    const json = JSON.stringify(settings);
    if (items.length) {
      await list.items.getById(items[0].Id).update({ [CONFIG_JSON_FIELD]: json });
    } else {
      await list.items.add({ [CONFIG_JSON_FIELD]: json });
    }
  }

  private toIssue(item: SPItem, progressLog: IProgressLogEntry[]): IIssue {
    const f = ISSUE_FIELDS;
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
      createdBy: (item[f.createdBy] as string) || "",
      reportDate: (item[f.reportDate] as string) || "",
      departmentBU: (item[f.departmentBU] as string) || "",
      region: (item[f.region] as string) || "",
      alreadyInContact: (item[f.alreadyInContact] as IIssue["alreadyInContact"]) || "No",
      deviationType: (item[f.deviationType] as string) || "",
      issueOrigin: (item[f.issueOrigin] as string) || "",
      additionalComments: (item[f.additionalComments] as string) || "",

      followUp: (item[f.followUp] as string) || "",
      taskOwner: (item[f.taskOwner] as string) || "",
      ownerBU: (item[f.ownerBU] as string) || (item[f.departmentBU] as string) || "",
      dueDate: (item[f.dueDate] as string) || "",

      rootCause: (item[f.rootCause] as string) || "",
      correctiveAction: (item[f.correctiveAction] as string) || "",
      implementationDate: (item[f.implementationDate] as string) || "",
      effectivenessCheck: (item[f.effectivenessCheck] as string) || "",
      verifiedBy: (item[f.verifiedBy] as string) || "",
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
  private toSPFields(issue: Partial<IIssue>): Record<string, unknown> {
    const f = ISSUE_FIELDS;
    const out: Record<string, unknown> = {};
    const set = <K extends keyof IIssue>(key: K, spName: string, transform?: (v: IIssue[K]) => unknown): void => {
      if (issue[key] === undefined) return;
      out[spName] = transform ? transform(issue[key] as IIssue[K]) : issue[key];
    };

    set("qsNumber", f.qsNumber);
    set("shortSummary", f.shortSummary);
    set("description", f.description);
    set("immediateAction", f.immediateAction);
    set("severity", f.severity);
    set("createdBy", f.createdBy);
    set("reportDate", f.reportDate);
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

    set("taskOwner", f.taskOwner);
    set("ownerBU", f.ownerBU);
    set("dueDate", f.dueDate);

    set("rootCause", f.rootCause);
    set("correctiveAction", f.correctiveAction);
    set("implementationDate", f.implementationDate);
    set("effectivenessCheck", f.effectivenessCheck);
    set("verifiedBy", f.verifiedBy);
    set("verifiedDate", f.verifiedDate);
    set("closedDate", f.closedDate);
    set("closedAt", f.closedAt);

    set("holdReason", f.holdReason);
    set("holdUntil", f.holdUntil);

    set("ownerUpdate", f.ownerUpdate, boolToYesNo);
    set("ownerUpdateAt", f.ownerUpdateAt);
    set("ownerUpdateText", f.ownerUpdateText);

    return out;
  }
}
