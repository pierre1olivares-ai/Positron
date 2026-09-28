import assert from "node:assert/strict";
import test from "node:test";

import type { IIssue } from "../src/webparts/qstarIssueManager/models/IIssue";
import {
  buildIssueTransition,
  effectivenessTestEnd,
  NC_TEST_STATUS,
  reopenIssuePatch,
  startEffectivenessTestPatch,
} from "../src/webparts/qstarIssueManager/domain/issueLifecycle";

const now = new Date(2026, 8, 28, 12, 30, 0);

function issue(overrides: Partial<IIssue> = {}): IIssue {
  return {
    id: 1, qsNumber: 1001, triaged: true, status: "In Progress", taskCreated: "Yes", transformedInto: "NC Minor",
    shortSummary: "Cold-chain breach", description: "Logger recorded an excursion", immediateAction: "Quarantine", severity: "High",
    createdBy: "Reporter", reportDate: "2026-06-01", departmentBU: "Quality", region: "Western Europe (Amsterdam)", alreadyInContact: "Yes",
    deviationType: "Quality", issueOrigin: "Internal Finding", additionalComments: "", followUp: "", taskOwner: "Owner", ownerBU: "Quality", dueDate: "2026-07-01",
    rootCause: "", correctiveAction: "", implementationDate: "", effectivenessCheck: "", verifiedBy: "", verifiedDate: "", closedDate: "", closedAt: "",
    holdReason: "", holdUntil: "", ownerUpdate: false, ownerUpdateAt: "", ownerUpdateText: "", reminderCycle: "initial", attachments: [], progressLog: [],
    ...overrides,
  };
}

test("closing through a status-only patch cannot bypass the NC test gate", () => {
  assert.throws(() => buildIssueTransition(issue(), { status: "Closed" }, now), /Start the 2-month/);
  assert.throws(() => buildIssueTransition(issue({ implementationDate: "2026-01-01" }), { status: "Closed" }, now), /Start the 2-month/);
});

test("NC closure is blocked until the test's calendar end date", () => {
  const nc = issue({ status: NC_TEST_STATUS, implementationDate: "2026-07-29", verifiedBy: "Verifier", verifiedById: 42 });
  assert.throws(() => buildIssueTransition(nc, { status: "Closed" }, now), /2026-09-29/);
  const endDate = new Date(2026, 8, 29, 0, 0, 0);
  const closed = buildIssueTransition(nc, { status: "Closed" }, endDate);
  assert.equal(closed.closedDate, "2026-09-29");
  assert.equal(closed.closedAt, endDate.toISOString());
  assert.equal(closed.verifiedDate, "2026-09-29");
});

test("NC closure requires a verifier identity after the test period", () => {
  const nc = issue({ status: NC_TEST_STATUS, implementationDate: "2026-07-28" });
  assert.throws(() => buildIssueTransition(nc, { status: "Closed" }, now), /verifier/);
  assert.throws(() => buildIssueTransition(nc, { status: "Closed", verifiedBy: "Unresolved name" }, now), /identity/);
  const closed = buildIssueTransition(nc, { status: "Closed", verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com" }, now);
  assert.equal(closed.closedDate, "2026-09-28");
  assert.equal(closed.verifiedDate, "2026-09-28");
});

test("OFI closure succeeds without NC-only verification fields", () => {
  const closed = buildIssueTransition(issue({ transformedInto: "OFI" }), { status: "Closed" }, now);
  assert.equal(closed.status, "Closed");
  assert.equal(closed.closedDate, "2026-09-28");
  assert.equal(closed.closedAt, now.toISOString());
  assert.equal(closed.verifiedDate, undefined);
});

test("effectiveness start produces one patch with status and local implementation date", () => {
  assert.deepEqual(startEffectivenessTestPatch(issue(), now), { status: NC_TEST_STATUS, implementationDate: "2026-09-28" });
  assert.deepEqual(buildIssueTransition(issue(), { status: NC_TEST_STATUS }, now), { status: NC_TEST_STATUS, implementationDate: "2026-09-28" });
  assert.throws(() => startEffectivenessTestPatch(issue({ transformedInto: "OFI" }), now), /Only a nonconformity/);
  assert.throws(() => startEffectivenessTestPatch(issue({ implementationDate: "2026-09-29" }), now), /today or earlier/);
});

test("test windows ending in a shorter month use that month's final day", () => {
  const nc = issue({ status: NC_TEST_STATUS, implementationDate: "2026-12-31", verifiedBy: "Verifier", verifiedById: 42 });
  assert.equal(effectivenessTestEnd(nc), "2027-02-28");
  assert.equal(buildIssueTransition(nc, { status: "Closed" }, new Date(2027, 1, 28, 12)).closedDate, "2027-02-28");
});

test("putting an issue on hold requires a reason and a current or future date", () => {
  assert.throws(() => buildIssueTransition(issue(), { status: "On Hold" }, now), /reason/);
  assert.throws(() => buildIssueTransition(issue(), { status: "On Hold", holdReason: "Supplier dependency" }, now), /resume date/);
  assert.throws(() => buildIssueTransition(issue(), { status: "On Hold", holdReason: "Supplier dependency", holdUntil: "2026-09-27" }, now), /today or later/);
  const patch = buildIssueTransition(issue(), { status: "On Hold", holdReason: "Supplier dependency", holdUntil: "2026-09-28T00:00:00Z" }, now);
  assert.equal(patch.holdUntil, "2026-09-28");
});

test("an expired hold still allows follow-up and resuming work", () => {
  const held = issue({ status: "On Hold", holdReason: "Supplier dependency", holdUntil: "2026-09-27" });
  assert.deepEqual(buildIssueTransition(held, { followUp: "Checking supplier today" }, now), { followUp: "Checking supplier today" });
  assert.deepEqual(buildIssueTransition(held, { status: "In Progress" }, now), { status: "In Progress" });
});

test("reopening clears prior closure and effectiveness metadata", () => {
  const closed = issue({ status: "Closed", closedDate: "2026-09-27", closedAt: "2026-09-27T12:00:00Z", verifiedBy: "Verifier", verifiedById: 42, verifiedByEmail: "verifier@example.com", verifiedDate: "2026-09-27", implementationDate: "2026-07-27", effectivenessCheck: "Passed" });
  const patch = reopenIssuePatch(closed, now);
  assert.equal(patch.status, "In Progress");
  for (const field of ["closedDate", "closedAt", "verifiedBy", "verifiedByEmail", "verifiedDate", "implementationDate", "effectivenessCheck"] as const) {
    assert.equal(patch[field], "", field);
  }
  assert.equal(patch.verifiedById, 0);
  assert.equal(patch.reminderCycle, now.toISOString());
  assert.throws(() => buildIssueTransition(closed, { status: "Created" }, now), /Re-open/);
});

test("transition validation does not mutate the caller's issue or pending draft", () => {
  const original = issue();
  const patch = { status: NC_TEST_STATUS };
  buildIssueTransition(original, patch, now);
  assert.equal(original.status, "In Progress");
  assert.equal(original.implementationDate, "");
  assert.deepEqual(patch, { status: NC_TEST_STATUS });
});
