import assert from "node:assert/strict";
import { after, afterEach, beforeEach, test } from "node:test";
import type { IIssue } from "../src/webparts/qstarIssueManager/models/IIssue";
import { IssueConflictError, IssueRefreshError } from "../src/webparts/qstarIssueManager/services/issueErrors";
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
  default: App, QMIssueDetail, OwnerIssueDetail, TriageForm, ReporterForm, ProgressLog, SettingsView, Dashboard,
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

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: Error) => void;
  const promise = new Promise<T>((accept, fail) => { resolve = accept; reject = fail; });
  return { promise, resolve, reject };
}

function assertFieldsDisabled(disabled: boolean): void {
  const inputs = container.querySelectorAll<HTMLInputElement>("label input, label select, label textarea");
  assert.ok(inputs.length);
  for (const input of Array.from(inputs)) {
    if (!input.readOnly) assert.equal(input.disabled, disabled, input.closest("label")?.textContent || "Editable field");
  }
}

function fillReport(): void {
  change("Short summary *", "Temperature excursion");
  change("Description *", "Shipment arrived outside the allowed range");
  change("Severity *", "High");
  change("Department / Business Unit *", "Quality");
  change("Region *", "Western Europe (Amsterdam)");
  change("Deviation type *", "Quality");
  change("Where does it come from? *", "Internal Finding");
}

test("QM locks its draft during a save while the independent progress form stays editable", async () => {
  const pending = deferred<IIssue>();
  let writes = 0;
  render(QMIssueDetail, qmProps({ issue: issue({ verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com" }), onUpdate: () => { writes += 1; return pending.promise; } }));
  change("Follow up (Quality Team notes)", "Submitted note");
  await click("Save changes");
  assertFieldsDisabled(true);
  const progress = container.querySelector<HTMLTextAreaElement>("textarea[placeholder^='What did you do']");
  assert.ok(progress);
  assert.equal(progress.disabled, false);
  act(() => { Simulate.change(progress, { target: { value: "Independent progress draft" } } as never); button("Saving…").click(); });
  await act(async () => { pending.resolve(issue({ followUp: "Submitted note", eTag: '"2"' })); });
  assertFieldsDisabled(false);
  assert.equal(field("Follow up (Quality Team notes)").value, "Submitted note");
  assert.equal(progress.value, "Independent progress draft");
  assert.equal(button("Save changes").disabled, true);
  assert.equal(writes, 1);
});

for (const transformedInto of ["NC Minor", "OFI"] as const) {
  test(`owner ${transformedInto} form locks only its task fields during submission`, async () => {
    const pending = deferred<IIssue>();
    const original = issue({ transformedInto, status: "Created" });
    render(OwnerIssueDetail, { issue: original, owner: "Owner", onBack: () => undefined, onAddProgress: async () => undefined, onUpdate: () => pending.promise });
    change("Status", "In Progress");
    if (transformedInto === "NC Minor") change("Implementation date", todayDate());
    await click(transformedInto === "NC Minor" ? "Save progress" : "Save");
    assertFieldsDisabled(true);
    assert.equal(container.querySelector<HTMLTextAreaElement>("textarea")?.disabled, false);
    await act(async () => { pending.resolve({ ...original, status: "In Progress", implementationDate: todayDate(), eTag: '"2"' }); });
    assertFieldsDisabled(false);
    assert.equal(field("Status").value, "In Progress");
    if (transformedInto === "NC Minor") assert.equal(field("Implementation date").value, todayDate());
  });
}

test("triage locks assignment fields until a rejected request returns its draft", async () => {
  const pending = deferred<IIssue>();
  render(TriageForm, { issue: issue({ triaged: false }), onBack: () => undefined, onTriage: () => pending.promise });
  change("Task owner (gets reminders)", "New Owner");
  change("Task owner Microsoft 365 email", "new@example.com");
  change("Follow up note (optional)", "Submitted triage note");
  await click("Create issue");
  assertFieldsDisabled(true);
  assert.equal(button("Create issue").disabled, true);
  await act(async () => { pending.reject(new Error("Write rejected")); });
  assertFieldsDisabled(false);
  assert.equal(field("Follow up note (optional)").value, "Submitted triage note");
  assert.equal(button("Create issue").disabled, false);
});

test("reporter locks its form during creation and shows one accepted reference", async () => {
  const pending = deferred<number>();
  let creates = 0;
  render(ReporterForm, { reporterName: "Reporter", settings: {}, onSubmit: () => { creates += 1; return pending.promise; } });
  fillReport();
  await click("Submit report");
  assertFieldsDisabled(true);
  act(() => { button("Submitting…").click(); });
  await act(async () => { pending.resolve(3010); });
  assert.match(container.textContent || "", /Report submitted.*QS-3010/);
  assert.equal(creates, 1);
  await click("Report another");
  assertFieldsDisabled(false);
  assert.equal(field("Short summary *").value, "");
});

test("settings lock link edits during save while diagnostics remain available", async () => {
  const pending = deferred<void>();
  function SettingsHarness(): React.ReactElement {
    const [settings, setSettings] = React.useState({ msFormUrl: "", flowId: "" });
    return React.createElement(SettingsView, { settings, onRunDiagnostics: async () => [], onSave: async (next: typeof settings) => { await pending.promise; setSettings(next); } });
  }
  render(SettingsHarness, {});
  change("Q-Star Microsoft Form URL", "https://forms.office.com/new");
  await click("Save settings");
  assertFieldsDisabled(true);
  assert.equal(button("Run connection test").disabled, false);
  await act(async () => { pending.resolve(); });
  assertFieldsDisabled(false);
  assert.equal(field("Q-Star Microsoft Form URL").value, "https://forms.office.com/new");
  assert.equal(button("Save settings").disabled, true);
});

test("a pending append locks its text without locking the QM draft", async () => {
  const pending = deferred<void>();
  let appends = 0;
  render(QMIssueDetail, qmProps({ onUpdate: async () => issue(), onAddProgress: () => { appends += 1; return pending.promise; } }));
  const progress = container.querySelector<HTMLTextAreaElement>("textarea[placeholder^='What did you do']");
  assert.ok(progress);
  act(() => { Simulate.change(progress, { target: { value: "Completed mitigation" } } as never); });
  await click("Add update");
  assert.equal(progress.disabled, true);
  assertFieldsDisabled(false);
  change("Follow up (Quality Team notes)", "Independent QM draft");
  act(() => { button("Saving…").click(); });
  await act(async () => { pending.resolve(); });
  assert.equal(progress.disabled, false);
  assert.equal(progress.value, "");
  assert.equal(field("Follow up (Quality Team notes)").value, "Independent QM draft");
  assert.equal(appends, 1);
});

for (const component of [QMIssueDetail, OwnerIssueDetail]) {
  test(`${component.name} holds lock reason and date until persistence finishes`, async () => {
    const pending = deferred<IIssue>();
    render(component, { ...qmProps(), owner: "Owner", onUpdate: () => pending.promise });
    await click("Put on hold");
    change("Reason for hold *", "Waiting for parts");
    const until = addCalendarDays(todayDate(), 7);
    change("Resume work on *", until);
    const confirms = Array.from(container.querySelectorAll("button")).filter(candidate => candidate.textContent === "Put on hold");
    await act(async () => { Simulate.click(confirms[confirms.length - 1]); });
    assertFieldsDisabled(true);
    assert.equal(button("Cancel").disabled, true);
    await act(async () => { pending.reject(new Error("Write rejected")); });
    assertFieldsDisabled(false);
    assert.equal(field("Reason for hold *").value, "Waiting for parts");
    assert.equal(field("Resume work on *").value, until);
  });
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

for (const component of [QMIssueDetail, OwnerIssueDetail, TriageForm]) {
  test(`${component.name} treats accepted updates as saved and requires explicit reload`, async () => {
    let writes = 0;
    let reloads = 0;
    let exits = 0;
    const update = async () => { writes += 1; throw new IssueRefreshError(1); };
    render(component, {
      ...qmProps(), owner: "Owner", onUpdate: update, onTriage: update,
      onBack: () => { exits += 1; },
      onReload: async () => { reloads += 1; return issue({ followUp: "Accepted note", implementationDate: todayDate(), eTag: '"2"' }); },
    });
    let submit: string;
    if (component === QMIssueDetail) {
      change("Follow up (Quality Team notes)", "Accepted note");
      submit = "Save changes";
    } else if (component === OwnerIssueDetail) {
      change("Implementation date", todayDate());
      submit = "Save progress";
    } else {
      change("Task owner (gets reminders)", "Owner");
      change("Task owner Microsoft 365 email", "owner@example.com");
      submit = "Create issue";
    }
    await click(submit);
    assert.equal(container.querySelector('[role="alert"]'), null);
    assert.match(container.querySelector('[role="status"]')?.textContent || "", /changes were saved.*do not submit/i);
    assert.equal(button(submit).disabled, true);
    act(() => { button(submit).click(); });
    assert.equal(writes, 1);
    await click(component === TriageForm ? "Reload latest and return to queue" : "Reload latest and discard draft");
    assert.equal(reloads, 1);
    if (component === TriageForm) assert.equal(exits, 1);
    else assert.equal(container.querySelector('[role="status"]'), null);
  });
}

async function renderApp(overrides: object = {}, initial: IIssue[] = [issue()]): Promise<void> {
  const dataService = {
    loadIssues: async () => initial,
    loadSettings: async () => ({ msFormUrl: "", flowId: "", access: [] }),
    getIssue: async () => issue({ eTag: '"2"' }),
    updateIssue: async () => { throw new Error("Unexpected write"); },
    createIssue: async () => { throw new Error("Unexpected create"); },
    addProgressLogEntry: async () => { throw new Error("Unexpected append"); },
    saveSettings: async () => undefined,
    ...overrides,
  };
  await act(async () => {
    ReactDOM.render(React.createElement(App, { dataService, profile: "admin", userDisplayName: "Quality Manager", userEmail: "qm@example.com", developmentMode: false, onRunDiagnostics: async () => [] }), container);
  });
}

async function openIssue(): Promise<void> {
  const row = Array.from(container.querySelectorAll("button")).find(candidate => candidate.textContent?.includes("Cold-chain breach"));
  assert.ok(row);
  await act(async () => { Simulate.click(row); });
}

function assertSavedWarning(): void {
  assert.match(container.querySelector('[role="status"]')?.textContent || "", /Saved with a warning:/);
  assert.doesNotMatch(container.textContent || "", /failed and was not saved|Could not submit/);
  assert.equal(container.querySelector('[role="alert"]'), null);
}

test("accepted creation with failed readback shows its reference and reloads without another create", async () => {
  let creates = 0;
  let reads = 0;
  const saved = issue({ triaged: false, eTag: undefined, saveWarning: "The report was saved. Refresh to load its latest details; do not submit it again." });
  await renderApp({
    createIssue: async () => { creates += 1; return saved; },
    getIssue: async (id: number) => { assert.equal(id, saved.id); reads += 1; return { ...saved, eTag: '"2"', saveWarning: undefined }; },
  }, []);
  await click("Report");
  fillReport();
  await click("Submit report");
  assertSavedWarning();
  assert.match(container.textContent || "", /Report submitted.*QS-1001/);
  assert.equal(container.querySelector("textarea"), null);
  await click("Reload saved issue");
  assert.equal(reads, 1);
  assert.equal(creates, 1);
  assert.equal(container.querySelector('[role="status"]'), null);
});

test("accepted append clears submitted text and failed reload remains a saved warning", async () => {
  let appends = 0;
  let reads = 0;
  const entry = { id: 10, text: "Mitigation completed", author: "Owner", ts: "2026-09-28T12:00:00Z" };
  await renderApp({
    addProgressLogEntry: async () => { appends += 1; return { ...entry, saveWarning: "Your update was posted. Refresh to reload its server-recorded author and time; do not post it again." }; },
    getIssue: async () => { reads += 1; if (reads === 1) throw new Error("Read unavailable"); return issue({ progressLog: [entry] }); },
  });
  await click("Issue register");
  await openIssue();
  const input = container.querySelector<HTMLTextAreaElement>("textarea[placeholder^='What did you do']");
  assert.ok(input);
  act(() => { Simulate.change(input, { target: { value: entry.text } } as never); });
  await click("Add update");
  assertSavedWarning();
  assert.equal(input.value, "");
  assert.equal(button("Add update").disabled, true);
  await click("Reload saved issue");
  assertSavedWarning();
  assert.match(container.querySelector('[role="status"]')?.textContent || "", /reloading is still unavailable/);
  await click("Reload saved issue");
  assert.equal(container.querySelector('[role="status"]'), null);
  assert.equal(appends, 1);
  assert.equal(reads, 2);
  assert.equal(Array.from(container.querySelectorAll("p")).filter(element => element.textContent === entry.text).length, 1);
});

test("accepted update remains blocked after reopening until reload supplies a new version", async () => {
  let writes = 0;
  const versions: string[] = [];
  await renderApp({
    updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
      writes += 1;
      versions.push(eTag);
      if (writes === 1) throw new IssueRefreshError(1);
      return issue({ ...patch, eTag: '"3"' });
    },
    getIssue: async () => issue({ followUp: "Accepted note", eTag: '"2"' }),
  });
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Accepted note");
  await click("Save changes");
  assertSavedWarning();
  assert.equal(button("Save changes").disabled, true);
  await click("Back to register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Later note");
  assert.equal(button("Save changes").disabled, true);
  assert.equal(writes, 1);
  await click("Reload latest and discard draft");
  assert.equal(field("Follow up (Quality Team notes)").value, "Accepted note");
  assert.equal(container.querySelector('[role="status"]'), null);
  change("Follow up (Quality Team notes)", "Later note");
  await click("Save changes");
  assert.deepEqual(versions, ['"1"', '"2"']);
});

test("a rejected write retains the failure banner and editable draft", async () => {
  await renderApp({ updateIssue: async () => { throw new Error("Write rejected"); } });
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Unsubmitted note");
  await click("Save changes");
  assert.match(container.querySelector('[role="alert"]')?.textContent || "", /failed and was not saved: Write rejected/);
  assert.equal(container.querySelector('[role="status"]'), null);
  assert.equal(field("Follow up (Quality Team notes)").value, "Unsubmitted note");
  assert.equal(button("Save changes").disabled, false);
});

test("accepted reopening offers reload from the closed detail without resubmitting", async () => {
  let writes = 0;
  await renderApp({ updateIssue: async () => { writes += 1; throw new IssueRefreshError(1); } }, [issue({ status: "Closed" })]);
  await click("Issue register");
  await openIssue();
  await click("Re-open issue");
  assertSavedWarning();
  await click("Re-open issue");
  assert.equal(writes, 1);
  await click("Reload saved issue");
  assert.equal(container.querySelector('[role="status"]'), null);
  assert.equal(field("Status").value, "In Progress");
});
