import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, normalizeSettings } from "../models/ISettings";
import { normalizeDateOnly } from "../domain/calendarDates";
import { normalizeRegion } from "../domain/referenceData";
import { AcceptedWriteError, IssueConflictError, IssueRefreshError } from "./issueErrors";
import { backendError, backendJson, IBackendResponse, IBackendTransport } from "./BackendApiClient";
import { positiveId } from "./sharePointValues";

const CALENDAR_FIELDS: (keyof IIssue)[] = ["reportDate", "dueDate", "implementationDate", "verifiedDate", "closedDate", "holdUntil"];
const STRING_FIELDS: (keyof IIssue)[] = [
  "shortSummary", "description", "immediateAction", "createdBy", "departmentBU", "region", "deviationType", "issueOrigin", "additionalComments",
  "followUp", "taskOwner", "ownerBU", "rootCause", "correctiveAction", "effectivenessCheck", "verifiedBy", "closedAt", "holdReason", "ownerUpdateAt", "ownerUpdateText", "reminderCycle",
];

/** The API obeys the same optimistic-concurrency and accepted-write contract as direct SharePoint. */
export class BackendApiDataService implements IDataService {
  private readonly snapshots = new Map<number, IIssue>();

  constructor(private client: IBackendTransport) {}

  public async loadIssues(): Promise<IIssue[]> {
    const values = await backendJson(this.client, "/issues");
    if (!Array.isArray(values)) throw new Error("The backend returned an invalid issue list.");
    return values.map((value) => this.remember(readIssue(value)));
  }

  public async getIssue(id: number): Promise<IIssue> {
    const issue = readIssue(await backendJson(this.client, `/issues/${id}`));
    if (issue.id !== id) throw new Error("The backend returned a different issue. Reload the register.");
    return this.remember(issue);
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    const response = await this.client.request("/issues", "POST", writableIssue(issue));
    if (!response.ok) throw await backendError(response);
    let value: Partial<IIssue> | undefined;
    try {
      value = await response.json() as Partial<IIssue>;
      return this.remember(readIssue(value));
    } catch {
      const issueId = positiveId(value?.id) ? value.id : locationId(response, /^\/issues\/([1-9]\d*)\/?$/);
      const qsNumber = Number(response.headers.get("X-QStar-Reference"));
      throw new AcceptedWriteError("create", { issueId, qsNumber: positiveId(value?.qsNumber) ? value.qsNumber : positiveId(qsNumber) ? qsNumber : undefined });
    }
  }

  public async updateIssue(id: number, patch: Partial<IIssue>, expectedETag?: string): Promise<IIssue> {
    const previous = this.snapshots.get(id);
    const eTag = expectedETag || previous?.eTag;
    if (!eTag || eTag === "*") throw new Error("Refresh this issue before saving so its current version can be checked.");
    const changed: Partial<IIssue> = {};
    Object.keys(patch).forEach((key) => {
      const name = key as keyof IIssue;
      if (!previous || previous.eTag !== eTag || previous[name] !== patch[name]) (changed as Record<string, unknown>)[name] = patch[name];
    });
    const body = writableIssue(changed);
    if (!Object.keys(body).length) return this.getIssue(id);
    const response = await this.client.request(`/issues/${id}`, "PATCH", body, { "If-Match": eTag });
    if (response.status === 412) {
      let freshIssue: IIssue | undefined;
      try {
        const conflict = await response.json() as { meta?: { freshIssue?: unknown } };
        if (conflict.meta?.freshIssue) {
          const fresh = readIssue(conflict.meta.freshIssue);
          if (fresh.id === id) freshIssue = this.remember(fresh);
        }
      } catch { /* Preserve the conflict when its optional snapshot cannot be decoded. */ }
      if (!freshIssue) { try { freshIssue = await this.getIssue(id); } catch { /* A reload remains available in the UI. */ } }
      throw new IssueConflictError(id, freshIssue);
    }
    if (!response.ok) throw await backendError(response);
    this.snapshots.delete(id);
    try {
      const value = await response.json() as { saved?: boolean };
      if (value?.saved === true) throw new IssueRefreshError(id);
      const saved = readIssue(value);
      if (saved.id !== id || !saved.eTag) throw new IssueRefreshError(id);
      return this.remember(saved);
    } catch { throw new IssueRefreshError(id); }
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<IProgressLogEntry> {
    if (!entry.text.trim()) throw new Error("Enter a progress update before posting.");
    // The API records its authenticated caller and SharePoint Created time.
    const response = await this.client.request(`/issues/${id}/progress`, "POST", { text: entry.text.trim() });
    if (!response.ok) throw await backendError(response);
    let value: Partial<IProgressLogEntry> | undefined;
    try {
      value = await response.json() as Partial<IProgressLogEntry>;
      return readProgress(value);
    } catch {
      const locationEntryId = locationId(response, /^\/issues\/([1-9]\d*)\/progress\/([1-9]\d*)\/?$/, id);
      const entryHeader = response.headers.get("X-QStar-Entry-Id");
      const headerEntryId = entryHeader && /^[1-9]\d*$/.test(entryHeader) ? Number(entryHeader) : undefined;
      if ((response.headers.get("Location") && !locationEntryId) || (entryHeader && !positiveId(headerEntryId)) ||
          (locationEntryId && headerEntryId && locationEntryId !== headerEntryId)) throw new AcceptedWriteError("progress", { issueId: id });
      const entryId = locationEntryId || headerEntryId || (positiveId(value?.id) ? value.id : undefined);
      throw new AcceptedWriteError("progress", { issueId: id, entryId });
    }
  }

  public async loadSettings(): Promise<ISettings> {
    return normalizeSettings(await backendJson(this.client, "/settings"));
  }

  public async saveSettings(settings: ISettings): Promise<void> {
    const response = await this.client.request("/settings", "PUT", normalizeSettings(settings));
    if (!response.ok) throw await backendError(response);
  }

  private remember(issue: IIssue): IIssue {
    this.snapshots.set(issue.id, issue);
    return issue;
  }
}

function readIssue(value: unknown): IIssue {
  const issue = value as IIssue;
  if (!issue || !positiveId(issue.id) || !positiveId(issue.qsNumber)) throw new Error("The backend returned an invalid issue identity.");
  const result = { ...issue, eTag: typeof issue.eTag === "string" && issue.eTag !== "*" ? issue.eTag : undefined,
    status: issue.status || undefined, transformedInto: issue.transformedInto || undefined,
    taskCreated: issue.taskCreated || "No", severity: issue.severity || "Medium", alreadyInContact: issue.alreadyInContact || "No",
    triaged: !!issue.triaged, ownerUpdate: !!issue.ownerUpdate,
    attachments: Array.isArray(issue.attachments) ? issue.attachments : [],
    progressLog: Array.isArray(issue.progressLog) ? issue.progressLog.map(readProgress) : [],
  };
  STRING_FIELDS.forEach((key) => { (result as unknown as Record<string, unknown>)[key] = typeof issue[key] === "string" ? issue[key] : ""; });
  CALENDAR_FIELDS.forEach((key) => { (result as unknown as Record<string, unknown>)[key] = normalizeDateOnly(issue[key] as string); });
  result.region = normalizeRegion(result.region);
  return result;
}

function readProgress(value: unknown): IProgressLogEntry {
  const entry = value as IProgressLogEntry;
  if (!entry || !positiveId(entry.id) || typeof entry.text !== "string" || typeof entry.author !== "string" || typeof entry.ts !== "string") {
    throw new Error("The backend returned an invalid progress receipt.");
  }
  return { ...entry };
}

function writableIssue(issue: Partial<IIssue>): Partial<IIssue> {
  const result = { ...issue };
  ["id", "qsNumber", "eTag", "saveWarning", "progressLog", "attachments"].forEach((key) => { delete (result as Record<string, unknown>)[key]; });
  return result;
}

function locationId(response: IBackendResponse, pattern: RegExp, parentId?: number): number | undefined {
  const location = response.headers.get("Location");
  if (!location) return undefined;
  try {
    // The API may return a relative route or an absolute URL including /api/v1.
    const path = new URL(location, "https://receipt.invalid").pathname;
    const start = path.indexOf("/issues/");
    const match = pattern.exec(start >= 0 ? path.substring(start) : path);
    if (!match || (parentId !== undefined && Number(match[1]) !== parentId)) return undefined;
    const id = Number(match[parentId === undefined ? 1 : 2]);
    return positiveId(id) ? id : undefined;
  } catch { return undefined; }
}
