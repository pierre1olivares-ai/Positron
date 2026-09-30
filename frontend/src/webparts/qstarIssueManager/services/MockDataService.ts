import { IIssueHistory } from "../domain/issueHistory";
import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, DEFAULT_SETTINGS, normalizeSettings } from "../models/ISettings";
import { IssueConflictError } from "./issueErrors";

const STORAGE_KEY = "qstar:mock:issues:v1";
const SETTINGS_KEY = "qstar:mock:settings:v1";
const OFFSET_KEY = "qstar:mock:reference-offset:v1";

/**
 * In-browser localStorage data layer — the same persistence model as the
 * original prototype's `window.storage` shim. Used for local development
 * in the SharePoint Workbench (`gulp serve`) before a real "Q-Star Issues"
 * list is available, so UI work isn't blocked on tenant access.
 */
export class MockDataService implements IDataService {
  private readonly histories = new Map<number, IIssueHistory>();
  public async loadIssues(): Promise<IIssue[]> {
    return this.readIssues();
  }

  public async initializeIssues(issues: IIssue[], replace = false): Promise<IIssue[]> {
    const current = this.readIssues();
    if (current.length && !replace) return current;
    const initialized = issues.map((issue) => ({ ...blankIssue(), ...issue, eTag: "1" }));
    this.histories.clear();
    initialized.forEach(issue => this.recordVersion(issue));
    this.persist(initialized);
    window.localStorage.setItem(OFFSET_KEY, String(initialized.reduce((max, issue) => Math.max(max, issue.qsNumber), 1000)));
    return this.readIssues();
  }

  public async getIssue(id: number): Promise<IIssue> {
    const issue = this.readIssues().filter((candidate) => candidate.id === id)[0];
    if (!issue) throw new Error(`Issue ${id} was not found.`);
    return issue;
  }

  private readIssues(): IIssue[] {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? (JSON.parse(raw) as IIssue[]).map((issue) => ({ ...issue, eTag: issue.eTag || "1" })) : [];
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    // Keep each read-modify-write synchronous, like a single atomic storage operation.
    const all = this.readIssues();
    const nextId = all.reduce((max, i) => Math.max(max, i.id), 0) + 1;
    const offset = Number(window.localStorage.getItem(OFFSET_KEY)) || all.reduce((max, i) => Math.max(max, i.qsNumber || 0), 1000);
    window.localStorage.setItem(OFFSET_KEY, String(offset));
    const created: IIssue = {
      ...blankIssue(),
      ...issue,
      id: nextId,
      qsNumber: offset + nextId,
      eTag: "1",
    };
    this.recordVersion(created);
    all.push(created);
    this.persist(all);
    return created;
  }

  public async updateIssue(id: number, patch: Partial<IIssue>, expectedETag?: string): Promise<IIssue> {
    const all = this.readIssues();
    const previous = all.filter((issue) => issue.id === id)[0];
    if (!previous) throw new Error(`Issue ${id} was not found.`);
    if (expectedETag && expectedETag !== previous.eTag) throw new IssueConflictError(id, previous);
    const updated = { ...previous, ...patch, id, qsNumber: previous.qsNumber, eTag: String(Number(previous.eTag) + 1) };
    if (this.histories.get(id)?.eTag === previous.eTag) this.recordVersion(updated);
    else this.histories.delete(id);
    const next = all.map((i) => (i.id === id ? updated : i));
    this.persist(next);
    return updated;
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<IProgressLogEntry> {
    const all = this.readIssues();
    if (!all.some((issue) => issue.id === id)) throw new Error(`Issue ${id} was not found.`);
    const saved = { ...entry, ts: new Date().toISOString() };
    const next = all.map((i) =>
      i.id === id ? { ...i, progressLog: [...(i.progressLog || []), saved] } : i
    );
    this.persist(next);
    return saved;
  }

  public async getIssueHistory(id: number): Promise<IIssueHistory> {
    const issue = await this.getIssue(id);
    const history = this.histories.get(id);
    return history && history.eTag === issue.eTag ? JSON.parse(JSON.stringify(history)) : { issueId: id, current: {}, versions: [], complete: false };
  }

  private recordVersion(issue: IIssue): void {
    const previous = this.histories.get(issue.id);
    const now = new Date().toISOString();
    const label = `${issue.eTag}.0`;
    const fields = { Triaged: issue.triaged ? "Yes" : "No", TaskCreated: issue.taskCreated, Status: issue.status || "" };
    this.histories.set(issue.id, {
      issueId: issue.id, eTag: issue.eTag, complete: true,
      current: { Id: issue.id, Created: previous?.current.Created || now, Modified: now, OData__UIVersionString: label, ...fields },
      versions: [...(previous?.versions || []).map(row => ({ ...(row as object), IsCurrentVersion: false })),
        { VersionId: Number(issue.eTag) * 512, VersionLabel: label, Created: now, IsCurrentVersion: true, ...fields }],
    });
  }

  public async loadSettings(): Promise<ISettings> {
    const raw = window.localStorage.getItem(SETTINGS_KEY);
    return raw ? normalizeSettings(JSON.parse(raw)) : { ...DEFAULT_SETTINGS };
  }

  public async saveSettings(settings: ISettings): Promise<void> {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(normalizeSettings(settings)));
  }

  private persist(issues: IIssue[]): void {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(issues));
  }
}

function blankIssue(): IIssue {
  return {
    qsNumber: 0,
    id: 0,
    triaged: false,
    status: undefined,
    taskCreated: "No",
    transformedInto: undefined,
    shortSummary: "",
    description: "",
    immediateAction: "",
    severity: "Medium",
    createdBy: "",
    reportDate: "",
    departmentBU: "",
    region: "",
    alreadyInContact: "No",
    deviationType: "",
    issueOrigin: "",
    additionalComments: "",
    followUp: "",
    taskOwner: "",
    ownerBU: "",
    dueDate: "",
    rootCause: "",
    correctiveAction: "",
    implementationDate: "",
    effectivenessCheck: "",
    verifiedBy: "",
    verifiedDate: "",
    closedDate: "",
    closedAt: "",
    holdReason: "",
    holdUntil: "",
    ownerUpdate: false,
    ownerUpdateAt: "",
    ownerUpdateText: "",
    reminderCycle: "",
    attachments: [],
    progressLog: [],
  };
}
