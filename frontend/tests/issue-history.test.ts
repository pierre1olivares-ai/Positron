import assert from "node:assert/strict";
import { test } from "node:test";
import { historicalBacklog, issueOpenBefore } from "../src/webparts/qstarIssueManager/domain/issueHistory";
import type { IIssue } from "../src/webparts/qstarIssueManager/models/IIssue";
import { retainedHistory } from "./history-fixtures";
import { sharePointHttpHarness } from "./sharepoint-http-harness";
const Module = require("node:module");
const originalLoad = Module._load;
Module._load = function (name: string, ...args: unknown[]) {
  return name === "@microsoft/sp-http" ? { AadHttpClient: { configurations: { v1: {} } } } : originalLoad.call(this, name, ...args);
};
const { BackendApiDataService } = require("../src/webparts/qstarIssueManager/services/BackendApiDataService");
Module._load = originalLoad;
import { MockDataService } from "../src/webparts/qstarIssueManager/services/MockDataService";

const issue = { id: 1, eTag: '"5"', triaged: true, taskCreated: "Yes", status: "In Progress", reportDate: "2000-01-01" } as IIssue;
const history = () => retainedHistory(issue, [
  { at: "2025-12-01T00:00:00Z", status: "In Progress" },
  { at: "2025-12-20T00:00:00Z", status: "Closed" },
  { at: "2026-01-05T00:00:00Z", status: "In Progress" },
  { at: "2026-02-10T00:00:00Z", status: "Closed" },
  { at: "2026-03-01T00:00:00Z", status: "In Progress" },
]);

test("retained versions preserve closed intervals across multiple reopen cycles", () => {
  const h = history();
  const at = (date: string) => historicalBacklog([issue], { 1: h }, new Date(date));
  assert.equal(at("2025-11-01"), 0);
  assert.equal(at("2025-12-15"), 1);
  assert.equal(at("2026-01-01"), 0);
  assert.equal(at("2026-02-01"), 1);
  assert.equal(at("2026-02-01")! - at("2026-01-01")!, 1);
  assert.equal(at("2026-03-01"), 0);
  assert.equal(at("2026-04-01"), 1);
  assert.equal(issueOpenBefore(issue, { ...h, versions: [...h.versions].reverse() }, new Date("2026-02-01")), true);
});

test("exact event boundaries are exclusive and preserve same-day closed intervals", () => {
  const h = retainedHistory(issue, [
    { at: "2026-01-15T08:00:00Z", status: "In Progress" },
    { at: "2026-01-15T09:00:00Z", status: "Closed" },
    { at: "2026-01-15T15:00:00Z", status: "In Progress" },
  ]);
  const at = (instant: string) => issueOpenBefore(issue, h, new Date(instant));
  assert.equal(at("2026-01-15T08:00:00Z"), false);
  assert.equal(at("2026-01-15T08:00:00.001Z"), true);
  assert.equal(at("2026-01-15T09:00:00Z"), true);
  assert.equal(at("2026-01-15T09:00:00.001Z"), false);
  assert.equal(at("2026-01-15T12:00:00Z"), false);
  assert.equal(at("2026-01-15T15:00:00Z"), false);
  assert.equal(at("2026-01-15T15:00:00.001Z"), true);
});

test("version order resolves multiple changes at the same native timestamp", () => {
  const h = retainedHistory(issue, [
    { at: "2026-01-01T00:00:00Z", status: "In Progress" },
    { at: "2026-01-15T09:00:00Z", status: "Closed" },
    { at: "2026-01-15T09:00:00Z", status: "In Progress" },
  ]);
  h.versions.reverse();
  assert.equal(issueOpenBefore(issue, h, new Date("2026-01-15T09:00:00.001Z")), true);
  assert.equal(issueOpenBefore(issue, h, new Date("2026-01-15T09:00:00Z")), true);
});

test("historical eligibility uses recorded triage, task creation and rejection", () => {
  const rejected = { ...issue, status: "Rejected", taskCreated: "No" } as IIssue;
  const h = retainedHistory(rejected, [
    { at: "2026-01-01T00:00:00Z", status: "", triaged: "No", taskCreated: "No" },
    { at: "2026-02-01T00:00:00Z", status: "Created", taskCreated: "No" },
    { at: "2026-03-01T00:00:00Z", status: "Created" },
    { at: "2026-04-01T00:00:00Z", status: "Rejected", taskCreated: "No" },
  ]);
  assert.deepEqual(["01", "02", "03", "04"].map(month => issueOpenBefore(rejected, h, new Date(`2026-${month}-15`))), [false, false, true, false]);
});

for (const defect of ["trimmed", "gap", "invalid fields", "invalid date"] as const) test(`${defect} history leaves earlier periods unknown but preserves a reliable suffix`, () => {
  const h = history();
  if (defect === "trimmed") h.versions = h.versions.slice(2);
  if (defect === "gap") h.versions.splice(1, 1);
  if (defect === "invalid fields") delete (h.versions[1] as any).Triaged;
  if (defect === "invalid date") (h.versions[1] as any).Created = "2025-02-30T00:00:00Z";
  assert.equal(issueOpenBefore(issue, h, new Date("2026-01-01")), undefined);
  assert.equal(issueOpenBefore(issue, h, new Date("2026-02-01")), true);
  assert.equal(issueOpenBefore(issue, h, new Date("2025-11-01")), false);
});

for (const defect of ["revision", "current fields", "current timestamp", "current label", "current marker", "duplicate", "unsupported label", "unsupported shape", "incomplete"] as const) test(`${defect} history cannot establish backlog`, () => {
  const h = history();
  if (defect === "revision") h.eTag = '"4"';
  if (defect === "current fields") h.current.TaskCreated = "No";
  if (defect === "current timestamp") h.current.Modified = "2026-03-02T00:00:00Z";
  if (defect === "current label") h.current.OData__UIVersionString = "6.0";
  if (defect === "current marker") (h.versions[0] as any).IsCurrentVersion = true;
  if (defect === "duplicate") h.versions.push(h.versions[0]);
  if (defect === "unsupported label") (h.versions[0] as any).VersionLabel = "1.1";
  if (defect === "unsupported shape") h.versions = [{ FieldValuesAsText: h.current }];
  if (defect === "incomplete") h.complete = false;
  assert.equal(historicalBacklog([issue], { 1: h }, new Date("2026-04-01")), undefined);
});

test("an unknown current register member prevents a numeric aggregate even when rejected today", () => {
  const rejected = { ...issue, id: 2, status: "Rejected", reportDate: "2099-01-01" } as IIssue;
  assert.equal(historicalBacklog([issue, rejected], { 1: history() }, new Date("2026-02-01")), undefined);
  assert.equal(historicalBacklog([], {}, new Date("2026-02-01")), 0);
});

for (const nextKey of ["__next", "odata.nextLink", "@odata.nextLink"]) test(`direct history exhausts ${nextKey} using the raw version collection and rereads the revision`, async () => {
  const h = history();
  const harness = sharePointHttpHarness([{ ...h.current, "odata.etag": issue.eTag }], nextKey === "@odata.nextLink");
  harness.state.readReply = request => {
    if (!request.path.endsWith("/versions")) return undefined;
    assert.equal(new URL(request.url).searchParams.get("$select"), "*");
    const second = new URL(request.url).searchParams.has("$skiptoken");
    return Response.json(nextKey === "__next" ? { d: { results: second ? h.versions.slice(2) : h.versions.slice(0, 2), ...(second ? {} : { [nextKey]: `${request.url}&$skiptoken=2` }) } }
      : { value: second ? h.versions.slice(2) : h.versions.slice(0, 2), ...(second ? {} : { [nextKey]: `${request.url}&$skiptoken=2` }) });
  };
  try {
    const result = await harness.service.getIssueHistory(1);
    assert.deepEqual(result.versions, h.versions);
    assert.equal(issueOpenBefore(issue, result, new Date("2026-01-01")), false);
    assert.deepEqual(harness.state.requests.map(r => r.method), ["GET", "GET", "GET", "GET"]);
    assert.equal(harness.state.requests.filter(r => r.path.endsWith("items(1)")).length, 2);
  } finally { harness.restore(); }
});

for (const defect of ["denied", "missing envelope", "cycle", "other site", "other list", "snapshot changed"]) test(`direct history ${defect} never reports a complete usable snapshot`, async () => {
  const h = history();
  const harness = sharePointHttpHarness([{ ...h.current, "odata.etag": issue.eTag }]);
  harness.state.readReply = request => {
    if (!request.path.endsWith("/versions")) return undefined;
    if (defect === "denied") return new Response("Denied", { status: 403 });
    if (defect === "missing envelope") return Response.json(h.versions);
    if (defect === "snapshot changed") {
      harness.state.issues[0]["odata.etag"] = '"6"';
      return Response.json({ value: h.versions });
    }
    return Response.json({ value: h.versions, "@odata.nextLink": defect === "cycle" ? request.url : defect === "other site" ? request.url.replace("tenant.sharepoint.com", "evil.example.com") : request.url.replace("Q-Star%20Issues", "Another%20List").replace("Q-Star Issues", "Another List") });
  };
  try {
    if (defect === "snapshot changed") assert.equal((await harness.service.getIssueHistory(1)).complete, false);
    else await assert.rejects(harness.service.getIssueHistory(1));
    assert.ok(harness.state.requests.length <= 3);
    assert.ok(harness.state.requests.every(r => r.method === "GET" && r.url.startsWith("https://tenant.sharepoint.com/sites/quality/")));
  } finally { harness.restore(); }
});

test("backend history retains native fields and rejects mismatched issue identities", async () => {
  const calls: string[] = [];
  let value = history();
  const service = new BackendApiDataService({ request: async (path: string, method?: string) => {
    calls.push(path); assert.ok(!method || method === "GET"); return Response.json(value);
  } } as any);
  assert.deepEqual(await service.getIssueHistory(1), value);
  value = { ...value, issueId: 2 };
  await assert.rejects(service.getIssueHistory(1));
  assert.deepEqual(calls, ["/issues/1/history", "/issues/1/history"]);
});

test("mock history records actual local create/update cycles and never fabricates legacy coverage", async () => {
  const values = new Map<string, string>();
  const original = globalThis.window;
  globalThis.window = { localStorage: { getItem: (key: string) => values.get(key), setItem: (key: string, value: string) => values.set(key, value) } } as any;
  try {
    const service = new MockDataService();
    const created = await service.createIssue({ reportDate: "2000-01-01" });
    const closed = await service.updateIssue(created.id, { triaged: true, taskCreated: "Yes", status: "Closed" }, created.eTag);
    const reopened = await service.updateIssue(created.id, { status: "In Progress" }, closed.eTag);
    const retained = await service.getIssueHistory(created.id);
    assert.equal(retained.versions.length, 3);
    assert.equal(issueOpenBefore(reopened, retained, new Date("2001-01-01")), false);
    assert.equal(issueOpenBefore(reopened, retained, new Date(Date.now() + 1000)), true);
    assert.equal((await new MockDataService().getIssueHistory(created.id)).complete, false);
    assert.equal((await service.getIssue(created.id)).reportDate, "2000-01-01");
  } finally { globalThis.window = original; }
});
