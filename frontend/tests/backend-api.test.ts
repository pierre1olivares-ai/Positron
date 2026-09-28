import assert from "node:assert/strict";
import test from "node:test";
import type { IBackendResponse, IBackendTransport } from "../src/webparts/qstarIssueManager/services/BackendApiClient";
import { AcceptedWriteError, IssueConflictError, IssueRefreshError } from "../src/webparts/qstarIssueManager/services/issueErrors";
import { REGIONS, normalizeRegion } from "../src/webparts/qstarIssueManager/domain/referenceData";
const Module = require("node:module");
const load = Module._load;
Module._load = function (name: string, ...args: unknown[]) {
  return name === "@microsoft/sp-http" ? { AadHttpClient: { configurations: { v1: { version: 1 } } } } : load.call(this, name, ...args);
};
const { BackendApiClient } = require("../src/webparts/qstarIssueManager/services/BackendApiClient");
const { BackendApiDataService } = require("../src/webparts/qstarIssueManager/services/BackendApiDataService");
const { BackendRoleResolver } = require("../src/webparts/qstarIssueManager/services/BackendRoleResolver");
const { BackendDiagnosticsService } = require("../src/webparts/qstarIssueManager/services/BackendDiagnosticsService");
Module._load = load;
const issue = (patch = {}) => ({ id: 1, qsNumber: 3001, eTag: '"1"', shortSummary: "Shipment", createdById: 4, taskOwnerId: 7,
  taskOwnerEmail: "alex@example.com", verifiedById: 8, reportDate: "2026-09-28T00:00:00Z", dueDate: "2026-10-01T00:00:00Z",
  closedAt: "2026-09-28T08:23:42Z", ownerUpdateAt: "2026-09-28T09:12:01Z", reminderCycle: "reopen-2", region: "Asia Pacific", progressLog: [], ...patch });
const session = () => ({ role: "qm", source: "backend", user: { userId: 7, displayName: "API Caller", email: "api@example.com" },
  connection: { siteUrl: "https://tenant.sharepoint.com/sites/backend", issuesListName: "Backend Issues", progressListName: "Backend Progress", betaAccessMode: false } });
function response(value: unknown, status = 200, headers: Record<string, string> = {}, malformed = false): IBackendResponse {
  return { ok: status >= 200 && status < 300, status, statusText: "", headers: { get: key => headers[key] || null },
    json: async () => { if (malformed) throw new Error("Invalid JSON"); return value; } };
}
function harness(respond: (path: string, method: string, body?: unknown) => IBackendResponse) {
  const calls: { path: string; method: string; body?: unknown; headers?: Record<string, string> }[] = [];
  const transport: IBackendTransport = { request: async (path, method = "GET", body, headers) => {
    calls.push({ path, method, body, headers }); return respond(path, method, body);
  } };
  return { service: new BackendApiDataService(transport), transport, calls };
}

test("AAD transport sends the configured resource, HTTPS URL, JSON and If-Match", async () => {
  const calls: any[] = [];
  const context = { aadHttpClientFactory: { getClient: async (resource: string) => {
    calls.push(resource); return { fetch: async (...args: unknown[]) => { calls.push(args); return response(issue()); } };
  } } };
  await new BackendApiClient(context, "app-guid", "https://api.example.com/api/v1/").request("/issues/1", "PATCH", { followUp: "New note" }, { "If-Match": '"1"' });
  assert.equal(calls[0], "app-guid"); assert.equal(calls[1][0], "https://api.example.com/api/v1/issues/1");
  assert.equal(calls[1][2].headers["If-Match"], '"1"'); assert.deepEqual(JSON.parse(calls[1][2].body), { followUp: "New note" });
  await assert.rejects(new BackendApiClient(context, "app-guid", "http://api.example.com").request("/me"), /HTTPS/);
  await assert.rejects(new BackendApiClient(context, "", "https://api.example.com").request("/me"), /Entra/);
  assert.equal(calls.length, 2);
});

test("list and single reads preserve identity/version/time and normalize calendar dates and old regions", async () => {
  const h = harness(path => response(path === "/issues" ? [issue()] : issue()));
  const [loaded] = await h.service.loadIssues();
  assert.equal(loaded.eTag, '"1"'); assert.equal(loaded.createdById, 4); assert.equal(loaded.taskOwnerEmail, "alex@example.com"); assert.equal(loaded.verifiedById, 8);
  assert.equal(loaded.reportDate, "2026-09-28"); assert.equal(loaded.dueDate, "2026-10-01");
  assert.equal(loaded.closedAt, "2026-09-28T08:23:42Z"); assert.equal(loaded.ownerUpdateAt, "2026-09-28T09:12:01Z");
  assert.equal(loaded.reminderCycle, "reopen-2"); assert.equal(loaded.region, "Asia Pacific (Bangkok)");
  assert.equal((await h.service.getIssue(1)).id, 1); assert.deepEqual(h.calls.map(call => call.path), ["/issues", "/issues/1"]);
});

test("PATCH sends changed writable fields and exact version then returns the saved issue", async () => {
  const h = harness((_path, method, body) => response(method === "PATCH" ? issue({ ...body as object, eTag: '"2"' }) : issue()));
  const initial = await h.service.getIssue(1);
  const saved = await h.service.updateIssue(1, { ...initial, followUp: "Edited", qsNumber: 9999 }, initial.eTag);
  assert.equal(saved.eTag, '"2"'); assert.equal(saved.followUp, "Edited");
  assert.deepEqual(h.calls[1].body, { followUp: "Edited" }); assert.equal(h.calls[1].headers?.["If-Match"], '"1"');
  await assert.rejects(h.service.updateIssue(1, { followUp: "unsafe" }, "*"), /Refresh/);
  const fresh = harness(() => response(issue()));
  await assert.rejects(fresh.service.updateIssue(1, { followUp: "unsafe" }), /Refresh/); assert.equal(fresh.calls.length, 0);
});

test("412 becomes a typed conflict and reloads without repeating PATCH", async () => {
  const h = harness((_path, method) => method === "PATCH" ? response({ code: "ISSUE_CONFLICT" }, 412) : response(issue({ eTag: '"2"' })));
  await assert.rejects(h.service.updateIssue(1, { followUp: "Edited" }, '"1"'), (error: unknown) => {
    assert.ok(error instanceof IssueConflictError); assert.equal(error.freshIssue?.eTag, '"2"'); return true;
  });
  assert.deepEqual(h.calls.map(call => call.method), ["PATCH", "GET"]);
});

test("accepted PATCH receipt or malformed success remains a saved refresh error", async () => {
  for (const accepted of [response({ saved: true, issueId: 1, saveWarning: "Reload" }), response(undefined, 200, {}, true)]) {
    const h = harness(() => accepted);
    await assert.rejects(h.service.updateIssue(1, { followUp: "Edited" }, '"1"'), (error: unknown) => error instanceof IssueRefreshError && error.saved);
    assert.equal(h.calls.length, 1);
  }
});

test("accepted POST warnings/header fallbacks retain IDs without inventing server author/time", async () => {
  const h = harness(() => response(issue({ saveWarning: "Reference sync pending" }), 201));
  assert.equal((await h.service.createIssue({ shortSummary: "Submitted" })).saveWarning, "Reference sync pending"); assert.equal(h.calls.length, 1);
  const fallback = harness(path => path === "/issues" ? response(undefined, 201, { Location: "/api/v1/issues/21", "X-QStar-Reference": "3021" }, true)
    : response(undefined, 201, { Location: "https://api.example.com/api/v1/issues/1/progress/72" }, true));
  const created = await fallback.service.createIssue({ shortSummary: "Submitted" });
  assert.equal(created.id, 21); assert.equal(created.qsNumber, 3021); assert.equal(created.eTag, undefined); assert.ok(created.saveWarning);
  const note = await fallback.service.addProgressLogEntry(1, { text: "  Observation  ", author: "Forged", ts: "1900-01-01" });
  assert.equal(note.id, 72); assert.equal(note.ts, ""); assert.equal(note.author, ""); assert.ok(note.saveWarning);
  assert.deepEqual(fallback.calls[1].body, { text: "Observation" }); assert.equal(fallback.calls.length, 2);
});

test("known accepted POST with no readable identity is nonretryable", async () => {
  const h = harness(() => response(undefined, 201, {}, true));
  await assert.rejects(h.service.createIssue({ shortSummary: "Submitted" }), (error: unknown) => error instanceof AcceptedWriteError && error.saved && error.operation === "create");
  await assert.rejects(h.service.addProgressLogEntry(1, { text: "Observation", author: "A", ts: "" }), (error: unknown) => error instanceof AcceptedWriteError && error.operation === "progress");
  assert.equal(h.calls.length, 2);
});

test("settings failures remain visible and retired connection/role configuration is stripped", async () => {
  await assert.rejects(harness(() => response({ title: "Access denied" }, 403)).service.loadSettings(), /403.*Access denied/);
  const h = harness(() => response({ msFormUrl: "form", flowId: "flow", siteUrl: "wrong", access: [{ role: "admin" }] }));
  assert.deepEqual(await h.service.loadSettings(), { msFormUrl: "form", flowId: "flow", access: [] });
  await h.service.saveSettings({ msFormUrl: "form", flowId: "flow", access: [{ email: "x", name: "X", role: "admin" }] });
  assert.deepEqual(h.calls[1].body, { msFormUrl: "form", flowId: "flow", access: [] });
});

test("/me is authoritative and invalid/incomplete/disabled sessions fail closed", async () => {
  const h = harness(() => response(session()));
  const resolved = await new BackendRoleResolver(h.transport, "https://api.example.com/api/v1").resolve();
  assert.equal(resolved.role, "qm"); assert.equal(resolved.user.email, "api@example.com"); assert.equal(resolved.connection.issuesListName, "Backend Issues");
  for (const broken of [{ ...session(), role: "superadmin" }, { ...session(), user: undefined }, { ...session(), connection: {} }, { ...session(), source: "sharepoint" }]) {
    const bad = harness(() => response(broken));
    await assert.rejects(new BackendRoleResolver(bad.transport, "https://api.example.com").resolve(), /Access is disabled/); assert.equal(bad.calls.length, 1);
  }
  const disabled = harness(() => response({ title: "Backend disabled" }, 503));
  await assert.rejects(new BackendRoleResolver(disabled.transport, "https://api.example.com").resolve(), /503/);
});

test("diagnostics validate returned checks and canonical regions preserve legacy aliases", async () => {
  const h = harness(() => response({ wrong: true }));
  assert.equal((await new BackendDiagnosticsService(h.transport).run())[0].status, "fail");
  assert.ok(REGIONS.includes("France (Paris)")); assert.ok(REGIONS.includes("Asia Pacific (Bangkok)")); assert.ok(!REGIONS.includes("Asia Pacific"));
  assert.equal(normalizeRegion("Asia Pacific"), "Asia Pacific (Bangkok)"); assert.equal(normalizeRegion("Germany"), "Western Europe (Amsterdam)");
});

test("progress header-only identity is recoverable and inconsistent receipts remain quarantined", async () => {
  const good = harness(() => response(undefined, 201, { "X-QStar-Entry-Id": "72" }, true));
  assert.equal((await good.service.addProgressLogEntry(1, { text: "Saved", author: "", ts: "" })).id, 72);
  for (const headers of [
    { "X-QStar-Entry-Id": "72", Location: "/issues/1/progress/73" },
    { "X-QStar-Entry-Id": "72", Location: "/issues/2/progress/72" },
    { "X-QStar-Entry-Id": "invalid", Location: "/issues/1/progress/72" },
  ]) {
    const h = harness(() => response(undefined, 201, headers, true));
    await assert.rejects(h.service.addProgressLogEntry(1, { text: "Saved", author: "", ts: "" }), AcceptedWriteError);
    assert.equal(h.calls.length, 1);
  }
});
