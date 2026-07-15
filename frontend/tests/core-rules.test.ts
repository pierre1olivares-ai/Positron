import assert from "node:assert/strict";
import test from "node:test";

import {
  DEFAULT_ROLE_GROUP_NAMES,
  resolveBetaRoleFromSitePermissions,
  resolveRoleFromGroupTitles,
} from "../src/webparts/qstarIssueManager/services/roleRules";
import { readPersonValue } from "../src/webparts/qstarIssueManager/services/sharePointValues";
import {
  ISSUE_FIELD_NAMES,
  ISSUE_SELECT_FIELD_NAMES,
  PROGRESS_FIELD_NAMES,
} from "../src/webparts/qstarIssueManager/services/fieldMap";
import {
  buildDiagnosticIssuePayload,
  buildDiagnosticProgressPayload,
} from "../src/webparts/qstarIssueManager/services/diagnosticPayload";

test("role resolution uses highest privilege and is case-insensitive", () => {
  const result = resolveRoleFromGroupTitles([
    "q-star quality managers",
    "Q-STAR ADMINS",
  ]);

  assert.equal(result.role, "admin");
  assert.equal(result.matchedGroup, DEFAULT_ROLE_GROUP_NAMES.admin);
});

test("unknown authenticated users fail down to Reader", () => {
  assert.deepEqual(resolveRoleFromGroupTitles(["Site Members"]), { role: "reader" });
});

test("beta access maps existing site permissions without custom groups", () => {
  assert.equal(resolveBetaRoleFromSitePermissions({ manageWeb: true, editListItems: true }).role, "admin");
  assert.equal(resolveBetaRoleFromSitePermissions({ manageWeb: false, editListItems: true }).role, "qm");
  assert.equal(resolveBetaRoleFromSitePermissions({ manageWeb: false, editListItems: false }).role, "reader");
});

test("SharePoint Person values retain stable identity data", () => {
  assert.deepEqual(readPersonValue({
    Id: 42,
    Title: "Alex Owner",
    EMail: "alex.owner@example.com",
  }), {
    id: 42,
    displayName: "Alex Owner",
    email: "alex.owner@example.com",
  });
  assert.deepEqual(readPersonValue(null), { displayName: "", email: "" });
});

test("production field maps require native Person fields, not companion email columns", () => {
  for (const obsolete of ["ReportedByEmail", "TaskOwnerEmail", "VerifiedByEmail"]) {
    assert.equal(ISSUE_FIELD_NAMES.includes(obsolete), false);
  }
  assert.equal(PROGRESS_FIELD_NAMES.some((name) => name.endsWith("Email")), false);
  assert.equal(ISSUE_SELECT_FIELD_NAMES.includes("TaskOwner/Id"), true);
  assert.equal(ISSUE_SELECT_FIELD_NAMES.includes("TaskOwner/EMail"), true);
});

test("diagnostic payload satisfies required issue fields and uses a Person lookup ID", () => {
  const issue = buildDiagnosticIssuePayload();
  for (const required of ["Title", "ShortSummary", "Description", "QsNumber", "Severity", "DepartmentBU", "Region", "Triaged", "TaskCreated"]) {
    assert.notEqual(issue[required], undefined, `${required} should be populated`);
  }

  const progress = buildDiagnosticProgressPayload(12, 34);
  assert.equal(progress.ParentItemId, 12);
  assert.equal(progress.AuthorId, 34);
  assert.equal(typeof progress.EntryText, "string");
});
