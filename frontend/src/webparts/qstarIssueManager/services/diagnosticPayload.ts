import { ISSUE_FIELDS, PROGRESS_FIELDS } from "./fieldMap";
import { REGIONS } from "../domain/referenceData";

export const DIAGNOSTIC_MARKER = "Q-Star connection test — safe to delete";

export function buildDiagnosticIssuePayload(): Record<string, unknown> {
  return {
    Title: DIAGNOSTIC_MARKER,
    [ISSUE_FIELDS.shortSummary]: DIAGNOSTIC_MARKER,
    [ISSUE_FIELDS.description]: "Temporary item created by Q-Star connection diagnostics.",
    [ISSUE_FIELDS.qsNumber]: 0,
    [ISSUE_FIELDS.severity]: "Low",
    [ISSUE_FIELDS.reportDate]: new Date().toISOString(),
    [ISSUE_FIELDS.departmentBU]: "Quality",
    [ISSUE_FIELDS.region]: REGIONS[0],
    [ISSUE_FIELDS.triaged]: "No",
    [ISSUE_FIELDS.taskCreated]: "No",
  };
}

export function buildDiagnosticProgressPayload(issueId: number): Record<string, unknown> {
  return {
    Title: DIAGNOSTIC_MARKER,
    [PROGRESS_FIELDS.parentItemId]: issueId,
    [PROGRESS_FIELDS.entryDate]: new Date().toISOString(),
    [PROGRESS_FIELDS.text]: DIAGNOSTIC_MARKER,
  };
}
