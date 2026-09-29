import assert from "node:assert/strict";
import { after, afterEach, beforeEach, test } from "node:test";
import { retainedHistory } from "./history-fixtures";

const { JSDOM } = require("jsdom");
const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "http://localhost/" });
Object.assign(globalThis, { window: dom.window, document: dom.window.document, navigator: dom.window.navigator, MessageChannel: undefined });
dom.window.requestAnimationFrame = (callback: (timestamp: number) => void) => setTimeout(() => callback(Date.now()), 1);
dom.window.cancelAnimationFrame = clearTimeout;
Object.assign(globalThis, { requestAnimationFrame: dom.window.requestAnimationFrame, cancelAnimationFrame: clearTimeout });
const React = require("react") as typeof import("react");
const ReactDOM = require("react-dom") as typeof import("react-dom");
const { act, Simulate } = require("react-dom/test-utils") as typeof import("react-dom/test-utils");
const Module = require("node:module");
const originalLoad = Module._load;
Module._load = function (request: string, ...args: unknown[]): unknown {
  if (/\.(scss|css)$/.test(request)) return {};
  if (request === "recharts") return {
    ResponsiveContainer: ({ children }: { children: React.ReactNode }) => children,
    ComposedChart: ({ data }: { data: unknown }) => React.createElement("output", { "data-chart-input": true }, JSON.stringify(data)),
    BarChart: () => null,
    PieChart: () => null,
    Bar: () => null,
    XAxis: () => null,
    YAxis: () => null,
    Tooltip: () => null,
    CartesianGrid: () => null,
    Pie: () => null,
    Cell: () => null,
    Line: () => null,
    Legend: () => null,
    LabelList: () => null,
  };
  return originalLoad.call(this, request, ...args);
};
const { Dashboard } = require("../src/webparts/qstarIssueManager/components/QstarPrototype");
Module._load = originalLoad;

const RealDate = Date;
const originalTZ = process.env.TZ;
let container: HTMLDivElement;
let asOf: [number, number, number];
let hour: number;

beforeEach(() => {
  container = document.createElement("div");
  document.body.appendChild(container);
  asOf = [2026, 3, 15];
  hour = 12;
  globalThis.Date = new Proxy(RealDate, {
    construct: (target, args) => Reflect.construct(target, args.length ? args : [...asOf, hour]),
    get: (target, key) => key === "now" ? () => new RealDate(...asOf, hour).getTime() : Reflect.get(target, key),
  });
});

afterEach(() => {
  act(() => { ReactDOM.unmountComponentAtNode(container); });
  container.remove();
  globalThis.Date = RealDate;
  if (originalTZ === undefined) delete process.env.TZ;
  else process.env.TZ = originalTZ;
});

after(() => { dom.window.close(); });

const reported = (reportDate: string, transformedInto: string, closedDate = "") => ({
  reportDate, transformedInto, closedDate, triaged: true, taskCreated: "Yes", status: closedDate ? "Closed" : "In Progress",
});

const issues = [
  reported("2025-12-31", "OFI"),
  reported("2026-01-01", "NC Minor"),
  reported("2026-01-31T23:59:59Z", "OFI"),
  reported("2026-01-10", "Only sent to Dept/BU for Action", "2026-01-31T23:59:59Z"),
  reported("2026-01-31", "NC Major", "2026-02-01"),
  reported("2026-03-31", "NC Minor"),
  reported("2026-04-01", "OFI"),
  reported("2026-04-15T23:59:59Z", "Only sent to Dept/BU for Action"),
  reported("2026-04-16", "OFI"),
  reported("2025-12-01", "OFI", "2026-01-01"),
];

async function renderRecordedDashboard(records: any[]): Promise<void> {
  const issues = records.map((row, index) => ({ ...row, id: index + 1, eTag: `"${index + 1}"` }));
  const nativeTime = (date: string) => new Date(`${date.slice(0, 10)}T00:00:00`).toISOString();
  const histories = issues.map(row => retainedHistory(row, [
    { at: nativeTime(row.reportDate), status: row.closedDate ? "In Progress" : row.status, taskCreated: row.taskCreated },
    ...(row.closedDate ? [{ at: nativeTime(row.closedDate), status: "Closed", taskCreated: row.taskCreated }] : []),
  ]));
  await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues, dataService: { getIssueHistory: async (id: number) => histories[id - 1] } }), container); });
}

function chartInput(index: number): Record<string, string | number>[] {
  const output = container.querySelectorAll("[data-chart-input]")[index];
  assert.ok(output);
  return JSON.parse(output.textContent || "");
}

function click(text: string): void {
  const button = Array.from(container.querySelectorAll("button")).find(candidate => candidate.textContent === text);
  assert.ok(button);
  act(() => { Simulate.click(button); });
}

for (const timezone of ["Europe/Amsterdam", "America/New_York"]) {
  test(`native timestamps use local year and month boundaries in ${timezone}`, async () => {
    process.env.TZ = timezone;
    const row = { ...reported("2025-12-01", "OFI"), id: 1, eTag: '"3"' };
    const history = retainedHistory(row, [
      { at: "2025-12-01T00:00:00Z", status: "In Progress" },
      // These UTC instants are the next local day in Amsterdam, the same day in New York.
      { at: "2025-12-31T23:30:00Z", status: "Closed" },
      { at: "2026-01-31T23:30:00Z", status: "In Progress" },
    ]);
    await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: [row], dataService: { getIssueHistory: async () => history } }), container); });
    assert.deepEqual(chartInput(0).map(point => point.Backlog), timezone === "Europe/Amsterdam" ? [0, 1, 1, 1] : [1, 1, 1, 1]);
    assert.match(container.textContent || "", timezone === "Europe/Amsterdam" ? /Backlog flat year to date/ : /Backlog ▲ \+1 year to date/);
  });

  test(`month cutoffs remain local across daylight saving changes in ${timezone}`, async () => {
    process.env.TZ = timezone;
    const row = { ...reported("2025-12-01", "OFI", "2026-03-31"), id: 1, eTag: '"2"' };
    const history = retainedHistory(row, [
      { at: "2025-12-01T00:00:00Z", status: "In Progress" },
      // April 1 at 00:30 CEST, March 31 at 18:30 EDT.
      { at: "2026-03-31T22:30:00Z", status: "Closed" },
    ]);
    await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: [row], dataService: { getIssueHistory: async () => history } }), container); });
    assert.deepEqual(chartInput(0).map(point => point.Backlog), timezone === "Europe/Amsterdam" ? [1, 1, 1, 0] : [1, 1, 0, 0]);
  });

  test(`current backlog stops at the current instant through same-day close and reopen in ${timezone}`, async () => {
    process.env.TZ = timezone;
    asOf = [2026, 0, 15];
    // A later report-date edit must not erase the retained lifecycle history.
    const row = { ...reported("2099-01-01", "OFI"), id: 1, eTag: '"3"' };
    const history = retainedHistory(row, [
      { at: "2025-12-01T00:00:00Z", status: "In Progress" },
      { at: new RealDate(2026, 0, 15, 9).toISOString(), status: "Closed" },
      { at: new RealDate(2026, 0, 15, 15).toISOString(), status: "In Progress" },
    ]);
    const service = { getIssueHistory: async () => history };
    const render = async () => { await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: [row], dataService: service }), container); }); };
    await render();
    assert.equal(chartInput(0)[0].Backlog, 0);
    assert.equal(chartInput(0)[0].Created, 0);
    assert.match(container.textContent || "", /Backlog ▼ -1 year to date/);
    hour = 9;
    await render();
    assert.equal(chartInput(0)[0].Backlog, 0, "the current instant includes the close recorded exactly now");
    hour = 15;
    await render();
    assert.equal(chartInput(0)[0].Backlog, 1, "the current instant includes the reopen recorded exactly now");
    hour = 18;
    await render();
    assert.equal(chartInput(0)[0].Backlog, 1);
    assert.match(container.textContent || "", /Backlog flat year to date/);
  });

  test(`dashboard calendar boundaries and inclusive cutoffs in ${timezone}`, async () => {
    process.env.TZ = timezone;
    await renderRecordedDashboard(issues);
    assert.deepEqual(chartInput(0), [
      { name: "Jan 26", Created: 4, Closed: 2, Backlog: 4, Net: 2 },
      { name: "Feb 26", Created: 0, Closed: 1, Backlog: 3, Net: -1 },
      { name: "Mar 26", Created: 1, Closed: 0, Backlog: 4, Net: 1 },
      { name: "Apr 26", Created: 2, Closed: 0, Backlog: 6, Net: 2 },
    ]);
    assert.deepEqual(chartInput(1), [
      { name: "Dec '25", NC: 0, OFI: 2, Other: 0, Total: 2 },
      { name: "Jan '26", NC: 2, OFI: 1, Other: 1, Total: 4 },
      { name: "Feb '26", NC: 2, OFI: 1, Other: 1, Total: 4 },
      { name: "Mar '26", NC: 3, OFI: 1, Other: 1, Total: 5 },
      { name: "Apr '26", NC: 3, OFI: 2, Other: 2, Total: 7 },
    ]);
    const notes = Array.from(container.querySelectorAll("p")).map(element => element.textContent);
    assert.ok(notes.includes("Data calculated from 01 Jan 2026 to 15 Apr 2026."));
    assert.ok(notes.includes("Data calculated from 01 Dec 2025 to 15 Apr 2026."));
    assert.doesNotMatch(container.textContent || "", /Invalid Date/);
    const yearTotal = Array.from(container.querySelectorAll("div")).find(element => element.textContent === "2026 · YTD");
    assert.equal(yearTotal?.nextElementSibling?.textContent, "7");
    assert.match(container.textContent || "", /Backlog ▲ \+4 year to date/);
    click("Quarter");
    assert.deepEqual(chartInput(1), [
      { name: "Q4 '25", NC: 0, OFI: 2, Other: 0, Total: 2 },
      { name: "Q1 '26", NC: 3, OFI: 1, Other: 1, Total: 5 },
      { name: "Q2 '26", NC: 3, OFI: 2, Other: 2, Total: 7 },
    ]);
    click("Year");
    assert.deepEqual(chartInput(1), [
      { name: "2025", NC: 0, OFI: 2, Other: 0, Total: 2 },
      { name: "2026", NC: 3, OFI: 2, Other: 2, Total: 7 },
    ]);
    click("NC");
    assert.deepEqual(chartInput(1), [
      { name: "2025", NC: 0, Total: 0 },
      { name: "2026", NC: 3, Total: 3 },
    ]);
  });

  test(`January activity changes YTD backlog from its opening balance in ${timezone}`, async () => {
    process.env.TZ = timezone;
    asOf = [2026, 0, 31];
    const januaryIssues = [reported("2026-01-01", "OFI"), reported("2026-01-31T23:59:59Z", "NC Minor")];
    await renderRecordedDashboard(januaryIssues);
    assert.deepEqual(chartInput(0), [{ name: "Jan 26", Created: 2, Closed: 0, Backlog: 2, Net: 2 }]);
    assert.match(container.textContent || "", /Backlog ▲ \+2 year to date/);
  });

  test(`New Year's Day closures reduce carried backlog in ${timezone}`, async () => {
    process.env.TZ = timezone;
    asOf = [2026, 0, 1];
    const carriedIssues = [
      reported("2025-12-31T23:59:59Z", "OFI", "2026-01-01T23:59:59Z"),
      reported("2025-12-31", "NC Minor"),
      reported("2025-12-30", "OFI", "2025-12-31T23:59:59Z"),
      reported("2026-01-01", "OFI", "2026-01-01"),
      reported("2026-01-02", "OFI"),
    ];
    await renderRecordedDashboard(carriedIssues);
    assert.deepEqual(chartInput(0), [{ name: "Jan 26", Created: 1, Closed: 2, Backlog: 1, Net: -1 }]);
    assert.match(container.textContent || "", /Backlog ▼ -1 year to date/);
    await renderRecordedDashboard([]);
    assert.match(container.textContent || "", /Backlog flat year to date/);
  });

  test(`forwarded reports create no open tasks or backlog in ${timezone}`, async () => {
    process.env.TZ = timezone;
    asOf = [2026, 0, 1];
    const forwarded = { ...reported("2026-01-01", "Only sent to Dept/BU for Action"), status: "Created", taskCreated: "No" };
    await renderRecordedDashboard([forwarded]);
    assert.deepEqual(chartInput(0), [{ name: "Jan 26", Created: 1, Closed: 0, Backlog: 0, Net: 1 }]);
    const open = Array.from(container.querySelectorAll("div")).find(element => element.textContent === "Open issues");
    assert.equal(open?.nextElementSibling?.textContent, "0");
    assert.match(container.textContent || "", /Backlog flat year to date/);
    const carried = [
      forwarded, { ...forwarded, reportDate: "2025-12-30" },
      reported("2025-12-31", "OFI", "2026-01-01"),
      reported("2025-12-31", "NC Minor"),
    ];
    await renderRecordedDashboard(carried);
    assert.deepEqual(chartInput(0), [{ name: "Jan 26", Created: 1, Closed: 1, Backlog: 1, Net: 0 }]);
    assert.match(container.textContent || "", /Backlog ▼ -1 year to date/);
  });
}

test("reopening in January preserves December's closed baseline and reports +1", async () => {
  asOf = [2026, 0, 31];
  const row = { ...reported("2025-12-01", "OFI"), id: 1, eTag: '"3"' };
  const history = retainedHistory(row, [
    { at: "2025-12-01T00:00:00Z", status: "In Progress" },
    { at: "2025-12-20T00:00:00Z", status: "Closed" },
    { at: "2026-01-05T00:00:00Z", status: "In Progress" },
  ]);
  await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: [row], dataService: { getIssueHistory: async () => history } }), container); });
  assert.equal(chartInput(0)[0].Backlog, 1);
  assert.match(container.textContent || "", /Backlog ▲ \+1 year to date/);
  assert.match(container.textContent || "", /retained recorded history for the current register and current filters/);
});

test("history gaps remain null and unavailable while current rejected members stay in the historical cohort", async () => {
  const row = { ...reported("2025-12-01", "OFI"), id: 1, eTag: '"4"', status: "Rejected", taskCreated: "No", departmentBU: "Quality", region: "France (Paris)" };
  const h = retainedHistory(row, [
    { at: "2025-12-01T00:00:00Z", status: "In Progress", version: 1 },
    { at: "2026-03-05T00:00:00Z", status: "In Progress", version: 3 },
    { at: "2026-04-05T00:00:00Z", status: "Rejected", taskCreated: "No", version: 4 },
  ]);
  const unknown = { ...row, id: 2, departmentBU: "IT", region: "Western Europe (Amsterdam)" };
  const service = { getIssueHistory: async (id: number) => { if (id === 2) throw new Error("History denied"); return h; } };
  await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: [row, unknown], dataService: service }), container); });
  assert.deepEqual(chartInput(0).map(row => row.Backlog), [null, null, null, null]);
  assert.match(container.textContent || "", /Backlog unavailable year to date/);
  const filters = container.querySelectorAll("select");
  await act(async () => { Simulate.change(filters[0], { target: { value: "Quality" } } as never); });
  assert.deepEqual(chartInput(0).map(row => row.Backlog), [null, null, 1, 0]);
  assert.deepEqual(Array.from(container.querySelectorAll('table[aria-label="Historical backlog"] td')).map(cell => cell.textContent), ["Unavailable", "Unavailable", "1", "0"]);
  assert.match(container.textContent || "", /Backlog unavailable year to date/);
  await act(async () => { Simulate.change(filters[0], { target: { value: "" } } as never); Simulate.change(filters[1], { target: { value: "France (Paris)" } } as never); });
  assert.deepEqual(chartInput(0).map(row => row.Backlog), [null, null, 1, 0]);
});

test("dashboard history is bounded, ignores obsolete responses and refreshes changed revisions", async () => {
  const rows = Array.from({ length: 7 }, (_, index) => ({ ...reported("2025-12-01", "OFI"), id: index + 1, eTag: '"1"' }));
  const calls: { id: number; resolve: (value: unknown) => void }[] = [];
  const service = { getIssueHistory: (id: number) => new Promise(resolve => calls.push({ id, resolve })) };
  await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: rows, dataService: service }), container); });
  assert.equal(calls.length, 4);
  assert.equal(chartInput(0)[0].Backlog, null);
  const newer = [{ ...rows[0], eTag: '"2"', status: "Closed", closedDate: "2025-12-20" }];
  await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: newer, dataService: service }), container); });
  assert.equal(calls.length, 4);
  await act(async () => {
    calls.slice(0, 4).forEach(call => call.resolve(retainedHistory(rows[call.id - 1], [{ at: "2025-12-01T00:00:00Z", status: "In Progress" }])));
  });
  assert.equal(calls.length, 5);
  assert.equal(chartInput(0)[0].Backlog, null);
  await act(async () => { calls[4].resolve(retainedHistory(newer[0], [
    { at: "2025-12-01T00:00:00Z", status: "In Progress" }, { at: "2025-12-20T00:00:00Z", status: "Closed" },
  ])); });
  assert.deepEqual(chartInput(0).map(row => row.Backlog), [0, 0, 0, 0]);
  await act(async () => { ReactDOM.render(React.createElement(Dashboard, { issues: [...newer], dataService: service }), container); });
  assert.equal(calls.length, 5);
});
