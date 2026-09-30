import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { IIssueHistory } from "../domain/issueHistory";
import { ISettings } from "../models/ISettings";

export interface IDataService {
  /** Local development only; production services do not expose seeding. */
  initializeIssues?(issues: IIssue[], replace?: boolean): Promise<IIssue[]>;
  loadIssues(): Promise<IIssue[]>;
  getIssue(id: number): Promise<IIssue>;
  getIssueHistory(id: number): Promise<IIssueHistory>;
  createIssue(issue: Partial<IIssue>): Promise<IIssue>;
  updateIssue(id: number, patch: Partial<IIssue>, expectedETag?: string): Promise<IIssue>;
  addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<IProgressLogEntry>;

  loadSettings(): Promise<ISettings>;
  saveSettings(settings: ISettings): Promise<void>;
}
