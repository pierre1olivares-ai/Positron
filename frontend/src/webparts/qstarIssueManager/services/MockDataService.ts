import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, DEFAULT_SETTINGS } from "../models/ISettings";

const STORAGE_KEY = "qstar:mock:issues:v1";
const SETTINGS_KEY = "qstar:mock:settings:v1";

/**
 * In-browser localStorage data layer — the same persistence model as the
 * original prototype's `window.storage` shim. Used for local development
 * in the SharePoint Workbench (`gulp serve`) before a real "Q-Star Issues"
 * list is available, so UI work isn't blocked on tenant access.
 */
export class MockDataService implements IDataService {
  public async loadIssues(): Promise<IIssue[]> {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? (JSON.parse(raw) as IIssue[]) : [];
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    const all = await this.loadIssues();
    const nextId = all.reduce((max, i) => Math.max(max, i.id), 0) + 1;
    const nextQs = all.reduce((max, i) => Math.max(max, i.qsNumber || 0), 1000) + 1;
    const created: IIssue = {
      ...blankIssue(),
      ...issue,
      id: nextId,
      qsNumber: issue.qsNumber ?? nextQs,
    };
    all.push(created);
    this.persist(all);
    return created;
  }

  public async updateIssue(id: number, patch: Partial<IIssue>): Promise<void> {
    const all = await this.loadIssues();
    const next = all.map((i) => (i.id === id ? { ...i, ...patch } : i));
    this.persist(next);
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<void> {
    const all = await this.loadIssues();
    const next = all.map((i) =>
      i.id === id ? { ...i, progressLog: [...(i.progressLog || []), entry] } : i
    );
    this.persist(next);
  }

  public async loadSettings(): Promise<ISettings> {
    const raw = window.localStorage.getItem(SETTINGS_KEY);
    return raw ? { ...DEFAULT_SETTINGS, ...JSON.parse(raw) } : DEFAULT_SETTINGS;
  }

  public async saveSettings(settings: ISettings): Promise<void> {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
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
    attachments: [],
    progressLog: [],
  };
}
