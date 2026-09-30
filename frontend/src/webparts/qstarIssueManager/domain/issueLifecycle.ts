import type { IIssue, IssueStatus } from "../models/IIssue";
import { addCalendarMonths, calendarDaysBetween, normalizeDateOnly, todayDate } from "./calendarDates";

export const NC_TEST_STATUS: IssueStatus = "Under Testing/Revision";
export const NC_TEST_MONTHS = 2;

type DateOnlyIssueField = "reportDate" | "dueDate" | "implementationDate" | "verifiedDate" | "closedDate" | "holdUntil";
const DATE_ONLY_FIELDS: DateOnlyIssueField[] = ["reportDate", "dueDate", "implementationDate", "verifiedDate", "closedDate", "holdUntil"];
const STATUSES: IssueStatus[] = ["Created", "In Progress", NC_TEST_STATUS, "On Hold", "Closed", "Rejected"];

export function isNonconformity(issue: Pick<IIssue, "transformedInto">): boolean {
  return issue.transformedInto === "NC Minor" || issue.transformedInto === "NC Major";
}

export function effectivenessTestEnd(issue: Pick<IIssue, "implementationDate">): string {
  return addCalendarMonths(issue.implementationDate, NC_TEST_MONTHS);
}

/**
 * All editing paths, including a status dropdown, use this transition boundary.
 * The caller persists the returned patch and applies it to its draft only after
 * a successful save. DateTime audit fields remain ISO event timestamps.
 */
export function buildIssueTransition(current: IIssue, proposedPatch: Partial<IIssue>, now: Date = new Date()): Partial<IIssue> {
  const patch = { ...proposedPatch };
  const today = todayDate(now);
  if (!today) throw new Error("A valid current date is required.");

  DATE_ONLY_FIELDS.forEach((field) => {
    const value = patch[field];
    if (value === undefined) return;
    const normalized = normalizeDateOnly(value);
    if (value && !normalized) throw new Error(`Enter a valid date for ${field}.`);
    patch[field] = normalized;
  });

  const next = { ...current, ...patch };
  if (next.status !== undefined && STATUSES.indexOf(next.status) < 0) throw new Error("Choose a valid issue status.");
  const statusChanged = next.status !== current.status;

  if (current.status === "Closed" && statusChanged) {
    if (next.status !== "In Progress") throw new Error("Re-open the issue before changing its status.");
    return {
      ...patch,
      closedDate: "",
      closedAt: "",
      verifiedBy: "",
      verifiedById: 0,
      verifiedByEmail: "",
      verifiedDate: "",
      implementationDate: "",
      effectivenessCheck: "",
      reminderCycle: now.toISOString(),
    };
  }

  if (next.status === NC_TEST_STATUS) {
    if (!isNonconformity(next)) throw new Error("Only a nonconformity can enter effectiveness testing.");
    const implementationDate = normalizeDateOnly(next.implementationDate || today);
    if (!implementationDate || calendarDaysBetween(implementationDate, today) < 0) {
      throw new Error("The implementation date must be today or earlier before effectiveness testing can begin.");
    }
    if (!next.implementationDate || patch.implementationDate !== undefined || statusChanged) patch.implementationDate = implementationDate;
  }

  if (next.status === "On Hold") {
    if (!next.holdReason || !next.holdReason.trim()) throw new Error("Enter a reason before putting the issue on hold.");
    const holdUntil = normalizeDateOnly(next.holdUntil);
    if (!holdUntil) throw new Error("Enter a resume date before putting the issue on hold.");
    const dateChanged = patch.holdUntil !== undefined && normalizeDateOnly(current.holdUntil) !== holdUntil;
    if ((statusChanged || dateChanged) && calendarDaysBetween(today, holdUntil) < 0) {
      throw new Error("The resume date must be today or later.");
    }
  }

  if (next.status === "Closed" && current.status !== "Closed") {
    if (isNonconformity(next)) {
      if (current.status !== NC_TEST_STATUS) throw new Error("Start the 2-month effectiveness test before closing this nonconformity.");
      const end = effectivenessTestEnd(next);
      if (!end || calendarDaysBetween(end, today) < 0) {
        throw new Error(end ? `The effectiveness test must run until ${end} before this nonconformity can be closed.` : "Record the implementation date before closing this nonconformity.");
      }
      const hasVerifierIdentity = (next.verifiedById !== undefined && next.verifiedById > 0) || !!(next.verifiedByEmail && next.verifiedByEmail.trim());
      if (!next.verifiedBy || !next.verifiedBy.trim() || !hasVerifierIdentity) {
        throw new Error("Record the verifier's name and Microsoft 365 identity before closing this nonconformity.");
      }
      patch.verifiedDate = normalizeDateOnly(next.verifiedDate) || today;
    }
    patch.closedDate = today;
    patch.closedAt = now.toISOString();
  }

  if (next.status === "Rejected") patch.taskCreated = "No";
  return patch;
}

export function startEffectivenessTestPatch(issue: IIssue, now: Date = new Date()): Partial<IIssue> {
  return buildIssueTransition(issue, { status: NC_TEST_STATUS, implementationDate: issue.implementationDate || todayDate(now) }, now);
}

export function reopenIssuePatch(issue: IIssue, now: Date = new Date()): Partial<IIssue> {
  if (issue.status !== "Closed") throw new Error("Only a closed issue can be re-opened.");
  return buildIssueTransition(issue, { status: "In Progress" }, now);
}
