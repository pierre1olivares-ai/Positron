import { IIssue } from "../models/IIssue";

const EDITABLE_FIELDS: (keyof IIssue)[] = [
  "triaged", "status", "taskCreated", "transformedInto", "shortSummary", "description",
  "immediateAction", "severity", "reportDate", "departmentBU", "region", "alreadyInContact",
  "deviationType", "issueOrigin", "additionalComments", "followUp", "taskOwner", "taskOwnerEmail",
  "ownerBU", "dueDate", "rootCause", "correctiveAction", "implementationDate", "effectivenessCheck",
  "verifiedBy", "verifiedByEmail", "verifiedDate", "closedDate", "closedAt", "holdReason", "holdUntil",
  "ownerUpdate", "ownerUpdateAt", "ownerUpdateText",
];

/** Only values actually edited since the form was opened may be sent to the server. */
export function changedIssueFields(baseline: IIssue, draft: IIssue): Partial<IIssue> {
  const patch: Record<string, unknown> = {};
  for (const key of EDITABLE_FIELDS) {
    if (draft[key] !== baseline[key]) patch[key] = draft[key];
  }
  for (const person of ["taskOwner", "verifiedBy"] as const) {
    const emailKey = `${person}Email` as "taskOwnerEmail" | "verifiedByEmail";
    if (draft[person] !== baseline[person] || draft[emailKey] !== baseline[emailKey]) {
      patch[person] = draft[person];
      patch[emailKey] = draft[emailKey] || "";
      // No stale lookup ID may override an edited identity.
      patch[`${person}Id`] = !draft[person] && !draft[emailKey] ? 0 : undefined;
    }
  }
  return patch as Partial<IIssue>;
}
