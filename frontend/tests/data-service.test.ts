import assert from "node:assert/strict";
import test from "node:test";
import type { SPFI } from "@pnp/sp";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { SharePointDataService } from "../src/webparts/qstarIssueManager/services/SharePointDataService";
import { MockDataService } from "../src/webparts/qstarIssueManager/services/MockDataService";
import { AcceptedWriteError, IssueConflictError, IssueRefreshError } from "../src/webparts/qstarIssueManager/services/issueErrors";
import { normalizeSettings } from "../src/webparts/qstarIssueManager/models/ISettings";
import { sharePointHttpHarness } from "./sharepoint-http-harness";

type Row = Record<string, any>;
const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value));
const failure = (status: number) => Object.assign(new Error(`HTTP ${status}`), { status });
const root = "/sites/quality/Lists/Progress";

/** A stateful PnP transport double: real service mapping, paging and errors execute. */
function harness(initial: Row[] = []) {
  const state = {
    issues: initial.map(clone), progress: [] as Row[], offset: 3000, nextId: 10,
    writes: [] as { id: number; fields: Row; eTag: string | undefined }[],
    creates: [] as Row[], ensures: [] as string[], selections: [] as string[][],
    folders: new Set<string>(), canCreateFolder: true,
    failReads: false, failReferenceWrite: false, appendCalls: [] as { fields: Row[]; path: string }[],
  };
  const users: Record<string, Row> = {
    "old@example.com": { Id: 42, Title: "Alex Smith", EMail: "old@example.com" },
    "new@example.com": { Id: 99, Title: "Alex Smith", EMail: "new@example.com" },
  };
  function apply(row: Row, fields: Row): void {
    Object.keys(fields).forEach((key) => {
      if (["TaskOwnerId", "ReportedById", "VerifiedById"].includes(key)) {
        row[key.slice(0, -2)] = fields[key] == null ? null : clone(Object.values(users).find((user) => user.Id === fields[key]) || { Id: fields[key] });
      } else row[key] = fields[key];
    });
  }
  function getRows(title: string): Row[] {
    return title === "Q-Star Issues" ? state.issues : title === "Q-Star Config" ? [{ Id: 1, ReferenceOffset: state.offset }] : state.progress;
  }
  function item(title: string, id: number) {
    const query: any = async () => {
      if (state.failReads && title === "Q-Star Issues") throw failure(503);
      const row = getRows(title).find((candidate) => candidate.Id === id);
      if (!row) throw failure(404);
      return clone(row);
    };
    query.select = (...fields: string[]) => { state.selections.push(fields); return query; };
    query.expand = () => query;
    query.update = async (fields: Row, eTag?: string) => {
      state.writes.push({ id, fields: clone(fields), eTag });
      if (state.failReferenceWrite && "QsNumber" in fields) throw failure(503);
      const row = getRows(title).find((candidate) => candidate.Id === id);
      if (!row) throw failure(404);
      if (eTag && eTag !== "*" && eTag !== row["odata.etag"]) throw failure(412);
      apply(row, fields);
      row["odata.etag"] = `"${Number(String(row["odata.etag"] || "0").replace(/"/g, "")) + 1}"`;
    };
    return query;
  }
  function items(title: string) {
    let filter = "";
    let limit = Infinity;
    const rows = () => getRows(title).filter((row) => {
      if (filter.includes("FSObjType eq 0") && row.FSObjType === 1) return false;
      const folder = /FileDirRef eq '((?:''|[^'])*)'/.exec(filter);
      return !folder || row.FileDirRef === folder[1].replace(/''/g, "'");
    });
    const query: any = async () => clone(rows().slice(0, limit));
    query.select = (...fields: string[]) => { state.selections.push(fields); return query; };
    query.expand = () => query;
    query.filter = (value: string) => { filter = value; return query; };
    query.top = (value: number) => { limit = value; return query; };
    query[Symbol.asyncIterator] = async function* () {
      const all = rows();
      // Deliberately return small pages to exercise iteration even with tiny fixtures.
      for (let i = 0; i < all.length; i += 2) yield clone(all.slice(i, i + 2));
    };
    query.getById = (id: number) => item(title, id);
    query.using = () => query;
    query.add = async (fields: Row) => {
      state.creates.push(clone(fields));
      const row = { Id: state.nextId++, "odata.etag": '"1"' };
      apply(row, fields);
      state.issues.push(row);
      return clone(row);
    };
    return query;
  }
  function list(title: string) {
    const query: any = async () => ({ RootFolder: { ServerRelativeUrl: root } });
    query.select = () => query;
    query.expand = () => query;
    query.items = items(title);
    query.using = () => query;
    query.rootFolder = { folders: { addUsingPath: async (path: string) => {
      if (!state.canCreateFolder) throw failure(403);
      state.folders.add(path);
    } } };
    query.addValidateUpdateItemUsingPath = async (fields: Row[], path: string) => {
      state.appendCalls.push({ fields: clone(fields), path });
      const row = {
        Id: 71, FileDirRef: new URL(path).pathname, FSObjType: 0,
        EntryText: fields.find((field) => field.FieldName === "EntryText")?.FieldValue,
        Author: { Id: 7, Title: "Signed-in owner", EMail: "signed-in@example.com" },
        Created: "2026-09-28T14:00:00Z", EntryDate: "1900-01-01T00:00:00Z",
      };
      state.progress.push(row);
      return [{ FieldName: "Id", FieldValue: "71", ItemId: 71, HasException: false }];
    };
    return query;
  }
  const web: any = async () => ({ Url: "https://tenant.sharepoint.com/sites/quality" });
  web.select = () => web;
  web.lists = { getByTitle: list };
  web.ensureUser = async (email: string) => { state.ensures.push(email); return clone(users[email]); };
  web.getFolderByServerRelativePath = (path: string) => {
    const query: any = async () => {
      if (!state.folders.has(path)) throw failure(404);
      return { Exists: true };
    };
    query.select = () => query;
    return query;
  };
  const client = { web } as SPFI;
  const service = new SharePointDataService({} as WebPartContext, undefined, undefined, undefined, client);
  return { state, service, anotherService: () => new SharePointDataService({} as WebPartContext, undefined, undefined, undefined, client) };
}

const issue = (id: number, extra: Row = {}): Row => ({
  Id: id, QsNumber: 1200 + id, ShortSummary: "Existing issue", "odata.etag": '"1"',
  TaskOwner: { Id: 42, Title: "Alex Smith", EMail: "old@example.com" }, ...extra,
});

test("reads every page, projects native Person Name and normalizes only calendar fields", async () => {
  const { service, state } = harness([issue(1, { DueDate: "2026-09-28T00:00:00Z", ClosedAt: "2026-09-28T15:20:00Z" }), issue(2), issue(3)]);
  const loaded = await service.loadIssues();
  assert.equal(loaded.length, 3);
  assert.equal(loaded[0].dueDate, "2026-09-28");
  assert.equal(loaded[0].closedAt, "2026-09-28T15:20:00Z");
  assert.equal(loaded[0].eTag, '"1"');
  assert.ok(state.selections.some((fields) => fields.includes("TaskOwner/Name")));
  assert.equal(state.selections.flat().some((name) => name.endsWith("/LoginName")), false);
});

test("email edits resolve the new person and only changed fields are written", async () => {
  const { service, state } = harness([issue(1)]);
  const [loaded] = await service.loadIssues();
  const saved = await service.updateIssue(1, { ...loaded, taskOwnerEmail: "new@example.com", followUp: "New note" }, loaded.eTag);
  assert.deepEqual(state.ensures, ["new@example.com"]);
  assert.deepEqual(state.writes[0], { id: 1, fields: { FollowUp: "New note", TaskOwnerId: 99 }, eTag: '"1"' });
  assert.equal(saved.taskOwnerId, 99);
  assert.equal(saved.eTag, '"2"');
});

test("clearing a person and date sends null; references cannot be overwritten", async () => {
  const { service, state } = harness([issue(1, { DueDate: "2026-09-28T00:00:00Z" })]);
  const [loaded] = await service.loadIssues();
  await service.updateIssue(1, { dueDate: "", taskOwner: "", taskOwnerEmail: "", taskOwnerId: 0, qsNumber: 99999 }, loaded.eTag);
  assert.deepEqual(state.writes[0].fields, { DueDate: null, TaskOwnerId: null });
});

test("stale editors receive the latest issue on conflict without overwriting another save", async () => {
  const { service, state, anotherService } = harness([issue(1)]);
  const second = anotherService();
  const [firstView] = await service.loadIssues();
  const [secondView] = await second.loadIssues();
  await service.updateIssue(1, { status: "In Progress" }, firstView.eTag);
  await assert.rejects(second.updateIssue(1, { followUp: "Other editor" }, secondView.eTag), (error: unknown) => {
    assert.ok(error instanceof IssueConflictError);
    assert.equal(error.freshIssue?.status, "In Progress");
    return true;
  });
  assert.equal(state.issues[0].FollowUp, undefined);
});

test("a completed save whose refresh fails is distinguished from an unsaved mutation", async () => {
  const { service, state } = harness([issue(1)]);
  const [loaded] = await service.loadIssues();
  state.failReads = true;
  await assert.rejects(service.updateIssue(1, { followUp: "Saved note" }, loaded.eTag), (error: unknown) => {
    assert.ok(error instanceof IssueRefreshError);
    assert.equal(error.saved, true);
    return true;
  });
  assert.equal(state.issues[0].FollowUp, "Saved note");
});

test("concurrent creates use distinct server IDs while preserving legacy references", async () => {
  const { service, state, anotherService } = harness([issue(1, { QsNumber: 2040 })]);
  const [a, b] = await Promise.all([service.createIssue({ shortSummary: "A", qsNumber: 2041 }), anotherService().createIssue({ shortSummary: "B", qsNumber: 2041 })]);
  assert.deepEqual([a.qsNumber, b.qsNumber], [3010, 3011]);
  assert.equal(state.issues[0].QsNumber, 2040);
  assert.ok(state.creates.every((fields) => !("QsNumber" in fields)));
});

test("post-create failures return the saved reference and never cause a second POST", async () => {
  const { service, state } = harness();
  state.failReads = true;
  state.failReferenceWrite = true;
  const saved = await service.createIssue({ shortSummary: "New issue" });
  assert.equal(saved.id, 10);
  assert.equal(saved.qsNumber, 3010);
  assert.ok(saved.saveWarning);
  assert.equal(state.creates.length, 1);
  state.failReads = false;
  assert.equal((await service.getIssue(10)).qsNumber, 3010);
});

test("journal association comes from the protected folder, never forged ParentItemId", async () => {
  const { service, state } = harness([issue(1), issue(2)]);
  state.progress.push(
    { Id: 1, ParentItemId: 2, FileDirRef: `${root}/issue-1`, EntryText: "Scoped to one", Created: "2026-09-28T10:00:00Z" },
    { Id: 2, ParentItemId: 1, FileDirRef: root, EntryText: "Loose legacy row", Created: "2026-09-28T10:00:00Z" },
  );
  const loaded = await service.loadIssues();
  assert.deepEqual(loaded[0].progressLog.map((entry) => entry.text), ["Scoped to one"]);
  assert.equal(loaded[1].progressLog.length, 0);
});

test("journal append uses the absolute issue folder and server author/time", async () => {
  const { service, state } = harness([issue(1)]);
  const saved = await service.addProgressLogEntry(1, { ts: "1900-01-01T00:00:00Z", author: "Forged", authorId: 99, text: "Progress" });
  assert.equal(state.appendCalls[0].path, `https://tenant.sharepoint.com${root}/issue-1`);
  assert.deepEqual(state.appendCalls[0].fields.map((field) => field.FieldName), ["Title", "EntryText"]);
  assert.equal(saved.author, "Signed-in owner");
  assert.equal(saved.ts, "2026-09-28T14:00:00Z");
});

test("an owner cannot create an unprepared journal folder and no append is attempted", async () => {
  const { service, state } = harness([issue(1)]);
  state.canCreateFolder = false;
  await assert.rejects(service.addProgressLogEntry(1, { ts: "", author: "Owner", text: "Keep this draft" }), /still being prepared/);
  assert.equal(state.appendCalls.length, 0);
});

test("mock seeding persists and parallel journal/status writes preserve both", async () => {
  const data = new Map<string, string>();
  const oldWindow = globalThis.window;
  Object.assign(globalThis, { window: { localStorage: { getItem: (key: string) => data.get(key), setItem: (key: string, value: string) => data.set(key, value) } } });
  try {
    const service = new MockDataService();
    const initial = await service.createIssue({ shortSummary: "Seed" });
    await service.initializeIssues([{ ...initial, id: 224, qsNumber: 1200 }], true);
    await Promise.all([
      service.addProgressLogEntry(224, { ts: "", author: "Owner", text: "Progress" }),
      service.updateIssue(224, { ownerUpdate: true }, "1"),
    ]);
    const saved = await service.getIssue(224);
    assert.equal(saved.progressLog.length, 1);
    assert.equal(saved.ownerUpdate, true);
    await assert.rejects(service.updateIssue(224, { followUp: "Stale" }, "1"), IssueConflictError);
  } finally {
    Object.assign(globalThis, { window: oldWindow });
  }
});

test("legacy settings cannot override the connection or role configuration", () => {
  assert.deepEqual(normalizeSettings({ msFormUrl: "https://forms.office.com/report", flowId: "flow", spSiteUrl: "https://old.example", spListName: "old", access: [{ email: "someone", role: "admin" }] }), {
    msFormUrl: "https://forms.office.com/report", flowId: "flow", access: [],
  });
});

for (const operation of ["create", "progress"] as const) {
  const submit = (service: SharePointDataService) => operation === "create"
    ? service.createIssue({ shortSummary: "Submitted" })
    : service.addProgressLogEntry(1, { text: "Submitted", author: "Unverified", ts: "1900-01-01" });
  test(`direct ${operation} classifies only confirmed 2xx unreadable bodies as accepted without another POST`, async () => {
    const bodies = operation === "create"
      ? ["null", "{}", "[]", "123", '{"Id":0}', '{"Id":-1}', '{"Id":1.5}', '{"Id":"21"}', '{"Id":9007199254740992}']
      : ["null", "{}", "[]", '{"value":{}}', '{"value":[null]}', '{"value":[{"ItemId":71}]}', '{"value":[{"FieldName":"Id","FieldValue":"0","HasException":false}]}', '{"value":[{"FieldName":"Id","ItemId":-1,"HasException":false}]}', '{"value":[{"FieldName":"Id","FieldValue":"not-an-id","HasException":false}]}'];
    const replies = [
      ...["malformed", "", "   ", ...bodies].map(body => () => new Response(body, { status: 201 })),
      () => new Response(null, { status: 204 }),
      () => Object.assign(new Response("{}", { status: 200 }), { json: async () => { throw new TypeError("Body stream failed"); } }),
    ];
    for (const reply of replies) {
      const h = sharePointHttpHarness();
      h.state.reply = reply;
      try {
        await assert.rejects(submit(h.service), (error: unknown) => {
          assert.ok(error instanceof AcceptedWriteError);
          assert.equal(error.operation, operation);
          assert.equal(error.saved, true);
          if (operation === "progress") assert.equal(error.identity.issueId, 1);
          return true;
        });
        const writes = h.state.requests.filter(call => call.method === "POST");
        assert.equal(writes.length, 1);
        assert.equal(h.state.requests.indexOf(writes[0]), h.state.requests.length - 1);
      } finally { h.restore(); }
    }
  });

  test(`direct ${operation} keeps network/SyntaxError and non-2xx failures ordinary with one POST`, async () => {
    for (const failure of [new TypeError("Network unavailable"), new SyntaxError("Transport failure"), 403, 500, 302]) {
      const h = sharePointHttpHarness();
      h.state.reply = () => { if (failure instanceof Error) throw failure; return new Response("Rejected", { status: failure }); };
      try {
        await assert.rejects(submit(h.service), (error: any) => {
          assert.equal(error instanceof AcceptedWriteError, false);
          if (failure instanceof Error) assert.equal(error, failure);
          else { assert.equal(error.isHttpRequestError, true); assert.equal(error.status, failure); }
          return true;
        });
        assert.equal(h.state.requests.filter(call => call.method === "POST").length, 1);
      } finally { h.restore(); }
    }
  });
}

test("direct create preserves a known ID when its allocated reference is unsupported", async () => {
  const h = sharePointHttpHarness();
  h.state.reply = () => new Response('{"Id":9007199254740991}', { status: 201 });
  try {
    await assert.rejects(h.service.createIssue({ shortSummary: "Submitted" }), (error: unknown) => {
      assert.ok(error instanceof AcceptedWriteError);
      assert.equal(error.identity.issueId, 9007199254740991);
      return true;
    });
    assert.equal(h.state.requests.filter(call => call.method === "POST").length, 1);
  } finally { h.restore(); }
});

test("direct AddValidateUpdate field rejection remains an ordinary failure", async () => {
  const h = sharePointHttpHarness();
  h.state.reply = () => new Response(JSON.stringify({ value: [{ FieldName: "EntryText", HasException: true, ErrorMessage: "Field rejected" }] }), { status: 200 });
  try {
    await assert.rejects(h.service.addProgressLogEntry(1, { text: "Rejected", author: "", ts: "" }), (error: unknown) => {
      assert.equal(error instanceof AcceptedWriteError, false);
      assert.match((error as Error).message, /Field rejected/);
      return true;
    });
    assert.equal(h.state.requests.filter(call => call.method === "POST").length, 1);
  } finally { h.restore(); }
});

test("direct append readback failure keeps native audit details unknown until a successful read", async () => {
  const h = sharePointHttpHarness([issue(1)]);
  h.state.failProgressRead = true;
  try {
    const saved = await h.service.addProgressLogEntry(1, { text: "Accepted", author: "Unverified", authorEmail: "unverified@example.com", authorId: 99, ts: "1900-01-01" });
    assert.equal(saved.id, 71);
    assert.equal(saved.author, ""); assert.equal(saved.ts, "");
    assert.equal(saved.authorId, undefined); assert.equal(saved.authorEmail, undefined);
    assert.ok(saved.saveWarning);
    h.state.failProgressRead = false;
    const fresh = await h.service.getIssue(1);
    assert.equal(fresh.progressLog[0].ts, "2026-09-28T14:00:00Z");
    assert.equal(fresh.progressLog[0].author, "Native author");
    assert.equal(h.state.requests.filter(call => call.method === "POST").length, 1);
  } finally { h.restore(); }
});
