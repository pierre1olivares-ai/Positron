import assert from "node:assert/strict";
import { after, afterEach, beforeEach, test } from "node:test";
import type { IIssue } from "../src/webparts/qstarIssueManager/models/IIssue";
import { IssueConflictError } from "../src/webparts/qstarIssueManager/services/issueErrors";
import { addCalendarDays, todayDate } from "../src/webparts/qstarIssueManager/domain/calendarDates";

const { JSDOM } = require("jsdom");
const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "http://localhost/" });
Object.defineProperty(globalThis, "window", { value: dom.window, configurable: true });
Object.defineProperty(globalThis, "document", { value: dom.window.document, configurable: true });
Object.defineProperty(globalThis, "navigator", { value: dom.window.navigator, configurable: true });
// React 17's browser scheduler must not leave a Node MessageChannel alive after tests.
Object.defineProperty(globalThis, "MessageChannel", { value: undefined, configurable: true });
dom.window.requestAnimationFrame = (callback: (timestamp: number) => void) => setTimeout(() => callback(Date.now()), 1);
dom.window.cancelAnimationFrame = clearTimeout;
Object.defineProperty(globalThis, "requestAnimationFrame", { value: dom.window.requestAnimationFrame, configurable: true });
Object.defineProperty(globalThis, "cancelAnimationFrame", { value: clearTimeout, configurable: true });
// jsdom has no layout engine; provide stable geometry for the real chart components.
Object.defineProperty(globalThis, "ResizeObserver", { value: class {
  public observe(): void { /* fixed test viewport */ }
  public unobserve(): void { /* fixed test viewport */ }
  public disconnect(): void { /* no live observation */ }
}, configurable: true });
dom.window.HTMLElement.prototype.getBoundingClientRect = () => ({ width: 1024, height: 600, left: 0, top: 0, right: 1024, bottom: 600, x: 0, y: 0, toJSON: () => ({}) });

const React = require("react") as typeof import("react");
const ReactDOM = require("react-dom") as typeof import("react-dom");
const { act, Simulate } = require("react-dom/test-utils") as typeof import("react-dom/test-utils");
// SPFx generates .scss.ts shims; skip styles before tsx resolves those build shims.
const Module = require("node:module");
const originalLoad = Module._load;
Module._load = function (request: string, ...args: unknown[]): unknown {
  return /\.(scss|css)$/.test(request) ? {} : originalLoad.call(this, request, ...args);
};
const {
  QMIssueDetail, OwnerIssueDetail, TriageForm, ProgressLog, SettingsView, Dashboard,
} = require("../src/webparts/qstarIssueManager/components/QstarPrototype");
Module._load = originalLoad;

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement("div");
  document.body.appendChild(container);
});

afterEach(() => {
  act(() => { ReactDOM.unmountComponentAtNode(container); });
  container.remove();
});

after(() => { dom.window.close(); });

function issue(overrides: Partial<IIssue> = {}): IIssue {
  return {
    id: 1, qsNumber: 1001, eTag: '"1"', triaged: true, status: "In Progress", taskCreated: "Yes", transformedInto: "NC Minor",
    shortSummary: "Cold-chain breach", description: "Logger recorded an excursion", immediateAction: "Quarantine", severity: "High",
    createdBy: "Reporter", reportDate: "2026-06-01", departmentBU: "Quality", region: "Western Europe (Amsterdam)", alreadyInContact: "Yes",
    deviationType: "Quality", issueOrigin: "Internal Finding", additionalComments: "", followUp: "", taskOwner: "Owner", taskOwnerId: 7,
    taskOwnerEmail: "owner@example.com", ownerBU: "Quality", dueDate: "2026-07-01", rootCause: "", correctiveAction: "", implementationDate: "",
    effectivenessCheck: "", verifiedBy: "", verifiedDate: "", closedDate: "", closedAt: "", holdReason: "", holdUntil: "", ownerUpdate: false,
    ownerUpdateAt: "", ownerUpdateText: "", reminderCycle: "initial", attachments: [], progressLog: [], ...overrides,
  };
}

function render(component: React.ElementType, props: object): void {
  act(() => { ReactDOM.render(React.createElement(component, props), container); });
}

function field(label: string): HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement {
  const wrapper = Array.from(container.querySelectorAll("label")).find(candidate => candidate.firstElementChild?.textContent === label);
  assert.ok(wrapper, `Field ${label} should be rendered`);
  const input = wrapper.querySelector<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>("input,select,textarea");
  assert.ok(input);
  return input;
}

function change(label: string, value: string): void {
  act(() => { Simulate.change(field(label), { target: { value } } as never); });
}

function button(text: string): HTMLButtonElement {
  const result = Array.from(container.querySelectorAll("button")).find(candidate => candidate.textContent === text);
  assert.ok(result, `Button ${text} should be rendered`);
  return result;
}

async function click(text: string): Promise<void> {
  const target = button(text);
  assert.equal(target.disabled, false, `Button ${text} should be enabled`);
  await act(async () => { Simulate.click(target); });
}

function qmProps(overrides: object = {}): object {
  return { issue: issue(), onBack: () => undefined, onAddProgress: async () => undefined, onReload: async () => issue(), author: "Quality Manager", ...overrides };
}

test("the status dropdown cannot close an NC before its effectiveness test", async () => {
  let calls = 0;
  render(QMIssueDetail, qmProps({ onUpdate: async () => { calls += 1; return issue(); } }));
  change("Status", "Closed");
  await click("Save changes");
  assert.equal(calls, 0);
  assert.match(container.querySelector('[role="alert"]')?.textContent || "", /2-month effectiveness test/);
  assert.equal(field("Status").value, "Closed", "rejected saves retain the user's draft");
});

test("Verify and close completes an OFI without inaccessible NC verifier fields", async () => {
  const original = issue({ transformedInto: "OFI" });
  let saved: Partial<IIssue> | undefined;
  render(QMIssueDetail, qmProps({ issue: original, onUpdate: async (_id: number, patch: Partial<IIssue>) => { saved = patch; return { ...original, ...patch }; } }));
  assert.equal(Array.from(container.querySelectorAll("label")).some(label => label.firstElementChild?.textContent === "Verified by"), false);
  await click("Verify & close");
  assert.equal(saved?.status, "Closed");
  assert.equal(saved?.closedDate, todayDate());
  assert.ok(saved?.closedAt);
  assert.equal(container.querySelector('[role="alert"]'), null);
});

test("starting an effectiveness test updates the draft before a subsequent note save", async () => {
  let stored = issue();
  const calls: { patch: Partial<IIssue>; eTag: string }[] = [];
  render(QMIssueDetail, qmProps({ issue: stored, onUpdate: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
    calls.push({ patch, eTag });
    stored = { ...stored, ...patch, eTag: `"${calls.length + 1}"` };
    return stored;
  } }));
  await click("Start 2-month effectiveness test");
  assert.equal(field("Status").value, "Under Testing/Revision");
  assert.equal(field("Implementation date").value, todayDate());
  assert.equal(button("Save changes").disabled, true);
  change("Follow up (Quality Team notes)", "Observe the next shipment");
  await click("Save changes");
  assert.deepEqual(calls[1], { patch: { followUp: "Observe the next shipment" }, eTag: '"2"' });
  assert.equal(stored.status, "Under Testing/Revision");
  assert.equal(stored.implementationDate, todayDate());
});

test("a conflict retains the QM draft until the explicit reload action", async () => {
  const latest = issue({ followUp: "Another manager's note", taskOwner: "New owner", taskOwnerEmail: "new@example.com", eTag: '"2"' });
  let reloads = 0;
  render(QMIssueDetail, qmProps({
    onUpdate: async () => { throw new IssueConflictError(1, latest); },
    onReload: async () => { reloads += 1; return latest; },
  }));
  change("Follow up (Quality Team notes)", "My unsaved note");
  await click("Save changes");
  assert.equal(field("Follow up (Quality Team notes)").value, "My unsaved note");
  assert.equal(button("Save changes").disabled, true);
  assert.match(container.querySelector('[role="alert"]')?.textContent || "", /draft has been kept/);
  await click("Reload latest and discard draft");
  assert.equal(reloads, 1);
  assert.equal(field("Follow up (Quality Team notes)").value, "Another manager's note");
  assert.equal(field("Task owner").value, "New owner");
  assert.equal(container.querySelector('[role="alert"]'), null);
});

test("a newly created fallback record can acquire its missing version before editing", async () => {
  const fallback = issue({ eTag: undefined });
  const latest = issue({ followUp: "Saved during intake", eTag: '"7"' });
  let reloads = 0;
  const writes: { patch: Partial<IIssue>; eTag: string }[] = [];
  render(QMIssueDetail, qmProps({
    issue: fallback,
    onReload: async (id: number) => { assert.equal(id, fallback.id); reloads += 1; return latest; },
    onUpdate: async (_id: number, patch: Partial<IIssue>, eTag: string) => { writes.push({ patch, eTag }); return { ...latest, ...patch, eTag: '"8"' }; },
  }));
  change("Follow up (Quality Team notes)", "Draft before recovery");
  assert.equal(button("Save changes").disabled, true);
  assert.equal(button("Start 2-month effectiveness test").disabled, true);
  assert.equal(writes.length, 0);
  await click("Reload latest and discard draft");
  assert.equal(reloads, 1);
  assert.equal(field("Follow up (Quality Team notes)").value, "Saved during intake");
  assert.equal(Array.from(container.querySelectorAll("button")).some(candidate => candidate.textContent === "Reload latest and discard draft"), false);
  change("Follow up (Quality Team notes)", "Ready after recovery");
  await click("Save changes");
  assert.deepEqual(writes, [{ patch: { followUp: "Ready after recovery" }, eTag: '"7"' }]);
});

test("a failed progress submission preserves text for a successful retry", async () => {
  let fail = true;
  const accepted: string[] = [];
  render(ProgressLog, { entries: [], canAdd: true, author: "Owner", onAdd: async (entry: { text: string }) => {
    if (fail) throw new Error("Connection unavailable");
    accepted.push(entry.text);
  } });
  const input = container.querySelector("textarea");
  assert.ok(input);
  act(() => { Simulate.change(input, { target: { value: "Mitigation completed; evidence attached elsewhere." } } as never); });
  await click("Add update");
  assert.equal(input.value, "Mitigation completed; evidence attached elsewhere.");
  assert.match(container.querySelector('[role="alert"]')?.textContent || "", /Connection unavailable/);
  fail = false;
  await click("Add update");
  assert.deepEqual(accepted, ["Mitigation completed; evidence attached elsewhere."]);
  assert.equal(input.value, "");
});

test("an owner can retry a failed test-start without losing the implementation date", async () => {
  let fail = true;
  let journalWrites = 0;
  let stored = issue();
  const date = addCalendarDays(todayDate(), -7);
  function OwnerHarness(): React.ReactElement {
    const [current, setCurrent] = React.useState(stored);
    return React.createElement(OwnerIssueDetail, {
      issue: current, owner: "Owner", onBack: () => undefined,
      onAddProgress: async () => { journalWrites += 1; },
      onUpdate: async (_id: number, patch: Partial<IIssue>) => {
        if (fail) throw new Error("Temporary write failure");
        stored = { ...stored, ...patch, eTag: '"2"' };
        setCurrent(stored);
        return stored;
      },
    });
  }
  render(OwnerHarness, {});
  change("Implementation date", date);
  await click("Mitigation implemented — start 2-month test");
  assert.equal(field("Implementation date").value, date);
  assert.match(container.querySelector('[role="alert"]')?.textContent || "", /Temporary write failure/);
  assert.equal(stored.status, "In Progress");
  assert.equal(journalWrites, 0);
  fail = false;
  await click("Mitigation implemented — start 2-month test");
  assert.equal(stored.status, "Under Testing/Revision");
  assert.equal(stored.implementationDate, date);
  assert.match(container.textContent || "", /Your mitigation is in its 2-month effectiveness test/);
});

test("a failed triage save retains assignment and follow-up fields", async () => {
  let exits = 0;
  render(TriageForm, { issue: issue({ triaged: false, status: undefined }), onBack: () => { exits += 1; }, onTriage: async () => { throw new Error("SharePoint unavailable"); } });
  change("Task owner (gets reminders)", "New Owner");
  change("Task owner Microsoft 365 email", "new@example.com");
  change("Follow up note (optional)", "Review with the customer");
  await click("Create issue");
  assert.equal(exits, 0);
  assert.equal(field("Task owner (gets reminders)").value, "New Owner");
  assert.equal(field("Task owner Microsoft 365 email").value, "new@example.com");
  assert.equal(field("Follow up note (optional)").value, "Review with the customer");
  assert.match(container.querySelector('[role="alert"]')?.textContent || "", /SharePoint unavailable/);
  assert.equal(button("Create issue").disabled, false);
});

test("triage offers reload before attempting to save an intake without a version", async () => {
  const actions: string[] = [];
  render(TriageForm, {
    issue: issue({ triaged: false, status: undefined, eTag: undefined }),
    onBack: () => { actions.push("queue"); },
    onReload: async () => { actions.push("reload"); return issue({ triaged: false, status: undefined, eTag: '"2"' }); },
    onTriage: async () => { actions.push("write"); },
  });
  assert.equal(button("Create issue").disabled, true);
  assert.equal(button("Reject (no action)").disabled, true);
  await click("Reload latest and return to queue");
  assert.deepEqual(actions, ["reload", "queue"]);
});

test("settings show the active web-part connection and preserve failed link edits", async () => {
  render(SettingsView, {
    settings: { msFormUrl: "https://forms.office.com/old", flowId: "", spSiteUrl: "https://stale.example", spListName: "Stale list" },
    connection: { siteUrl: "https://tenant.sharepoint.com/sites/active", issuesListName: "Active issues", progressListName: "Active progress", betaAccessMode: true },
    onSave: async () => { throw new Error("Settings save failed"); }, onRunDiagnostics: async () => [],
  });
  assert.match(container.textContent || "", /https:\/\/tenant.sharepoint.com\/sites\/active/);
  assert.match(container.textContent || "", /Active issues/);
  assert.match(container.textContent || "", /Active progress/);
  assert.doesNotMatch(container.textContent || "", /stale\.example|Stale list/);
  assert.deepEqual(Array.from(container.querySelectorAll("label")).map(label => label.firstElementChild?.textContent), ["Q-Star Microsoft Form URL", "Form responses → SharePoint flow ID (optional)"]);
  change("Q-Star Microsoft Form URL", "https://forms.office.com/new");
  await click("Save settings");
  assert.equal(field("Q-Star Microsoft Form URL").value, "https://forms.office.com/new");
  assert.equal(button("Save settings").disabled, false);
  assert.match(container.textContent || "", /Settings save failed/);
});

test("today's reported issue appears in both open totals and year-to-date category totals", () => {
  render(Dashboard, { issues: [issue({ transformedInto: "OFI", reportDate: new Date().toISOString() })] });
  const heading = Array.from(container.querySelectorAll("h3")).find(candidate => candidate.textContent === "Issue category mix");
  assert.ok(heading);
  assert.match(heading.parentElement?.textContent || "", /1 issues/);
  const open = Array.from(container.querySelectorAll("span")).find(candidate => candidate.textContent === "Open issues");
  assert.ok(open);
  assert.match(open.parentElement?.parentElement?.textContent || "", /Open issues11 created total/);
});
