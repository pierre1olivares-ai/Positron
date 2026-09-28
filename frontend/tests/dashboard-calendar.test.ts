import assert from "node:assert/strict";
import { after, afterEach, beforeEach, test } from "node:test";

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

beforeEach(() => {
  container = document.createElement("div");
  document.body.appendChild(container);
  globalThis.Date = new Proxy(RealDate, {
    construct: (target, args) => Reflect.construct(target, args.length ? args : [2026, 3, 15, 12]),
    get: (target, key) => key === "now" ? () => new RealDate(2026, 3, 15, 12).getTime() : Reflect.get(target, key),
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
  test(`dashboard calendar boundaries and inclusive cutoffs in ${timezone}`, () => {
    process.env.TZ = timezone;
    act(() => { ReactDOM.render(React.createElement(Dashboard, { issues }), container); });
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
}
