import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings } from "../models/ISettings";

export interface IDataService {
  loadIssues(): Promise<IIssue[]>;
  createIssue(issue: Partial<IIssue>): Promise<IIssue>;
  updateIssue(id: number, patch: Partial<IIssue>): Promise<void>;
  addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<void>;

  loadSettings(): Promise<ISettings>;
  saveSettings(settings: ISettings): Promise<void>;
}
