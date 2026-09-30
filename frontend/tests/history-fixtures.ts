import type { IIssue } from "../src/webparts/qstarIssueManager/models/IIssue";
import type { IIssueHistory } from "../src/webparts/qstarIssueManager/domain/issueHistory";

export type HistoryEvent = { at: string; status: string; triaged?: string; taskCreated?: string; version?: number };
export function retainedHistory(issue: Pick<IIssue, "id" | "eTag">, events: HistoryEvent[], nativeCreated = events[0].at): IIssueHistory {
  const versions = events.map((event, index) => ({
    VersionId: (event.version || index + 1) * 512, VersionLabel: `${event.version || index + 1}.0`, Created: event.at,
    IsCurrentVersion: index === events.length - 1, Triaged: event.triaged || "Yes", TaskCreated: event.taskCreated || "Yes", Status: event.status,
  }));
  const last = versions[versions.length - 1];
  return { issueId: issue.id, eTag: issue.eTag, complete: true, versions,
    current: { Id: issue.id, Created: nativeCreated, Modified: last.Created, OData__UIVersionString: last.VersionLabel, Triaged: last.Triaged, TaskCreated: last.TaskCreated, Status: last.Status } };
}
