import assert from "node:assert/strict";
import { after, afterEach, beforeEach, test } from "node:test";
import type { IIssue, IProgressLogEntry } from "../src/webparts/qstarIssueManager/models/IIssue";
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
  assert.equal(field(label).disabled, false, `Field ${label} should be enabled`);
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

test("QM locks its detail and progress forms during a save and retains both drafts", async () => {
  const pending = deferred<IIssue>();
  let writes = 0;
  render(QMIssueDetail, qmProps({ issue: issue({ verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com" }), onUpdate: () => { writes += 1; return pending.promise; } }));
  change("Follow up (Quality Team notes)", "Submitted note");
  writeProgress("Independent progress draft");
  await click("Save changes");
  assertFieldsDisabled(true);
  const progress = container.querySelector<HTMLTextAreaElement>("textarea[placeholder^='What did you do']");
  assert.ok(progress);
  assert.equal(progress.disabled, true);
  act(() => { button("Saving…").click(); });
  await act(async () => { pending.resolve(issue({ followUp: "Submitted note", eTag: '"2"' })); });
  assertFieldsDisabled(false);
  assert.equal(field("Follow up (Quality Team notes)").value, "Submitted note");
  assert.equal(progress.value, "Independent progress draft");
  assert.equal(button("Save changes").disabled, true);
  assert.equal(writes, 1);
});

for (const transformedInto of ["NC Minor", "OFI"] as const) {
  test(`owner ${transformedInto} form locks task fields and progress during submission`, async () => {
    const pending = deferred<IIssue>();
    const original = issue({ transformedInto, status: "Created" });
    render(OwnerIssueDetail, { issue: original, owner: "Owner", onBack: () => undefined, onAddProgress: async () => undefined, onUpdate: () => pending.promise });
    change("Status", "In Progress");
    if (transformedInto === "NC Minor") change("Implementation date", todayDate());
    await click(transformedInto === "NC Minor" ? "Save progress" : "Save");
    assertFieldsDisabled(true);
    assert.equal(container.querySelector<HTMLTextAreaElement>("textarea")?.disabled, true);
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

test("a pending append locks the issue while retaining its unsaved QM draft", async () => {
  const pending = deferred<IProgressLogEntry>();
  let appends = 0;
  await renderApp({ addProgressLogEntry: () => { appends += 1; return pending.promise; } });
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Independent QM draft");
  const progress = container.querySelector<HTMLTextAreaElement>("textarea[placeholder^='What did you do']");
  assert.ok(progress);
  act(() => { Simulate.change(progress, { target: { value: "Completed mitigation" } } as never); });
  await click("Add update");
  assert.equal(progress.disabled, true);
  assertFieldsDisabled(true);
  assert.equal(button("Save changes").disabled, true);
  act(() => { button("Saving…").click(); });
  await act(async () => { pending.resolve({ text: "Completed mitigation", author: "Quality Manager", ts: new Date().toISOString() }); });
  assert.equal(progress.disabled, false);
  assertFieldsDisabled(false);
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

async function renderApp(overrides: object = {}, initial: IIssue[] = [issue()], props: object = {}): Promise<void> {
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
    ReactDOM.render(React.createElement(App, { dataService, profile: "admin", userDisplayName: "Quality Manager", userEmail: "qm@example.com", developmentMode: false, onRunDiagnostics: async () => [], ...props }), container);
  });
}

async function openIssue(qsNumber?: number): Promise<void> {
  const row = Array.from(container.querySelectorAll("button")).find(candidate => candidate.textContent?.includes("Cold-chain breach") && (!qsNumber || candidate.textContent.includes(`QS-${qsNumber}`)));
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
  assert.equal(button("Re-open issue").disabled, true);
  act(() => { button("Re-open issue").click(); });
  assert.equal(writes, 1);
  await click("Reload saved issue");
  assert.equal(container.querySelector('[role="status"]'), null);
  assert.equal(field("Status").value, "In Progress");
});

function progressInput(): HTMLTextAreaElement {
  const input = container.querySelector<HTMLTextAreaElement>("textarea[placeholder^='What did you do']");
  assert.ok(input);
  return input;
}

function writeProgress(text: string): void {
  const input = progressInput();
  assert.equal(input.disabled, false);
  act(() => { Simulate.change(input, { target: { value: text } } as never); });
}

function journalText(): string[] {
  return Array.from(container.querySelectorAll("ol li p")).map(element => element.textContent || "");
}

for (const transformedInto of ["OFI", "NC Minor"] as const) {
  for (const action of ["Verify & close", "Save changes"]) {
    test(`${transformedInto} ${action} locks progress until the closed view replaces the form`, async () => {
      const original = issue({
        transformedInto, status: transformedInto === "OFI" ? "In Progress" : "Under Testing/Revision",
        implementationDate: addCalendarDays(todayDate(), -90), verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com",
      });
      const pending = deferred<void>();
      let writes = 0;
      await renderApp({ updateIssue: async (_id: number, patch: Partial<IIssue>) => {
        writes += 1;
        assert.equal(patch.status, "Closed");
        await pending.promise;
        return { ...original, ...patch, eTag: '"2"' };
      } }, [original]);
      await click("Issue register");
      await openIssue();
      assert.equal(progressInput().disabled, false);
      if (action === "Save changes") change("Status", "Closed");
      await click(action);
      assert.equal(progressInput().disabled, true);
      assert.equal(button("Add update").disabled, true);
      await act(async () => { pending.resolve(); });
      assert.equal(container.querySelector("textarea"), null);
      assert.equal(button("Re-open issue").disabled, false);
      assert.equal(writes, 1);
    });
  }
}

test("a rejected closing transition unlocks and preserves the existing progress draft", async () => {
  const pending = deferred<IIssue>();
  await renderApp({ updateIssue: () => pending.promise }, [issue({ transformedInto: "OFI" })]);
  await click("Issue register");
  await openIssue();
  writeProgress("Keep this progress draft");
  await click("Verify & close");
  assert.equal(progressInput().disabled, true);
  assert.equal(button("Add update").disabled, true);
  await act(async () => { pending.reject(new Error("Closure rejected")); });
  assert.equal(progressInput().disabled, false);
  assert.equal(progressInput().value, "Keep this progress draft");
  assert.equal(button("Add update").disabled, false);
});

for (const first of ["update", "append"] as const) {
  for (const withIds of [true, false]) {
    test(`${first} requested first preserves both identical journal entries ${withIds ? "with" : "without"} server IDs`, async () => {
      const initialEntry: IProgressLogEntry = { id: withIds ? 10 : undefined, text: "The same legitimate observation", author: "Quality Manager", ts: "2026-09-28T12:00:00Z" };
      let stored = issue({ progressLog: [initialEntry] });
      const gates = { update: deferred<void>(), append: deferred<void>() };
      const calls: string[] = [];
      const versions: string[] = [];
      await renderApp({
        updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
          calls.push("update");
          versions.push(eTag);
          stored = { ...stored, ...patch, eTag: '"2"' };
          const snapshot = { ...stored, progressLog: [...stored.progressLog] };
          await gates.update.promise;
          return snapshot;
        },
        addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
          calls.push("append");
          assert.equal(entry.text, initialEntry.text);
          const saved = { ...initialEntry, id: withIds ? 11 : undefined };
          stored = { ...stored, progressLog: [...stored.progressLog, saved] };
          await gates.append.promise;
          return saved;
        },
      }, [stored]);
      await click("Issue register");
      await openIssue();
      change("Follow up (Quality Team notes)", "Saved independently of the journal");
      writeProgress(initialEntry.text);
      const second = first === "update" ? "append" : "update";
      const update = button("Save changes");
      const append = button("Add update");
      await act(async () => {
        Simulate.click(first === "update" ? update : append);
        Simulate.click(second === "update" ? update : append);
      });
      assertFieldsDisabled(true);
      assert.equal(progressInput().disabled, true);
      await act(async () => { gates[second].resolve(); });
      await act(async () => { gates[first].resolve(); });
      assert.deepEqual(journalText(), [initialEntry.text, initialEntry.text]);
      assert.deepEqual(calls, [first, second]);
      assert.deepEqual(versions, ['"1"']);
      assert.equal(progressInput().value, "");
      assert.equal(field("Follow up (Quality Team notes)").value, "Saved independently of the journal");
      assert.equal(button("Save changes").disabled, true);
    });
  }
}

test("a queued append proceeds after an issue save fails without discarding either draft", async () => {
  const pending = deferred<IIssue>();
  const calls: string[] = [];
  await renderApp({
    updateIssue: () => { calls.push("update"); return pending.promise; },
    addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => { calls.push("append"); return { ...entry, id: 18 }; },
  });
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Retain this rejected draft");
  writeProgress("An independent accepted observation");
  await act(async () => { Simulate.click(button("Save changes")); Simulate.click(button("Add update")); });
  assert.deepEqual(calls, ["update"]);
  await act(async () => { pending.reject(new Error("Issue write rejected")); });
  assert.deepEqual(calls, ["update", "append"]);
  assert.deepEqual(journalText(), ["An independent accepted observation"]);
  assert.equal(field("Follow up (Quality Team notes)").value, "Retain this rejected draft");
  assert.equal(button("Save changes").disabled, false);
  assert.equal(progressInput().value, "");
  assert.match(container.textContent || "", /Issue write rejected/);
});

test("closing requires a fresh request after an accepted append settles", async () => {
  const pending = deferred<void>();
  let stored = issue({ transformedInto: "OFI" });
  const calls: string[] = [];
  await renderApp({
    addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
      calls.push("append");
      const saved = { ...entry, id: 20 };
      stored = { ...stored, progressLog: [saved] };
      await pending.promise;
      return saved;
    },
    updateIssue: async (_id: number, patch: Partial<IIssue>) => { calls.push("close"); stored = { ...stored, ...patch, eTag: '"2"' }; return stored; },
  }, [stored]);
  await click("Issue register");
  await openIssue();
  writeProgress("Final closure evidence");
  await click("Add update");
  assert.equal(button("Verify & close").disabled, true);
  await act(async () => { button("Verify & close").click(); });
  assert.deepEqual(calls, ["append"]);
  assert.equal(progressInput().disabled, true);
  await act(async () => { pending.resolve(); });
  assert.deepEqual(calls, ["append"]);
  assert.deepEqual(journalText(), ["Final closure evidence"]);
  assert.equal(progressInput().disabled, false);
  await click("Verify & close");
  assert.deepEqual(calls, ["append", "close"]);
  assert.deepEqual(journalText(), ["Final closure evidence"]);
  assert.equal(container.querySelector("textarea"), null);
});

for (const transformedInto of ["OFI", "NC Minor"] as const) {
  for (const action of ["Verify & close", "Save changes"]) {
    test(`${transformedInto} ${action} requires a fresh close after retrying a failed append`, async () => {
      let stored = issue({
        transformedInto, status: transformedInto === "OFI" ? "In Progress" : "Under Testing/Revision",
        implementationDate: addCalendarDays(todayDate(), -90), verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com",
      });
      const attempts = [deferred<void>(), deferred<void>()];
      let appends = 0;
      let closures = 0;
      await renderApp({
        addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
          const attempt = attempts[appends++];
          await attempt.promise;
          const saved = { ...entry, id: 21 };
          stored = { ...stored, progressLog: [...stored.progressLog, saved] };
          return saved;
        },
        updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
          closures += 1;
          assert.equal(patch.status, "Closed");
          assert.equal(eTag, '"1"');
          stored = { ...stored, ...patch, eTag: '"2"' };
          return stored;
        },
      }, [stored]);
      await click("Issue register");
      await openIssue();
      writeProgress("Closure evidence that must not be lost");
      if (action === "Save changes") change("Status", "Closed");
      await click("Add update");
      assert.equal(button(action).disabled, true);
      await act(async () => { button(action).click(); });
      assert.equal(closures, 0);
      assert.equal(progressInput().disabled, true);
      await act(async () => { attempts[0].reject(new Error("Journal append rejected")); });
      assert.equal(closures, 0);
      assert.notEqual(stored.status, "Closed");
      assert.equal(progressInput().value, "Closure evidence that must not be lost");
      assert.equal(progressInput().disabled, false);
      assert.equal(button("Add update").disabled, false);
      assert.match(container.querySelector('[role="alert"]')?.textContent || "", /Journal append rejected/);
      assert.deepEqual(journalText(), []);
      await click("Add update");
      assert.equal(button(action).disabled, true);
      await act(async () => { button(action).click(); });
      assert.equal(appends, 2);
      assert.equal(closures, 0);
      await act(async () => { attempts[1].resolve(); });
      assert.equal(closures, 0);
      assert.deepEqual(journalText(), ["Closure evidence that must not be lost"]);
      await click(action);
      assert.equal(closures, 1);
      assert.equal(stored.status, "Closed");
      assert.deepEqual(journalText(), ["Closure evidence that must not be lost"]);
      assert.equal(container.querySelector("textarea"), null);
      assert.equal(container.querySelector('[role="alert"]'), null);
    });
  }
}

for (const transformedInto of ["OFI", "NC Minor"] as const) {
  for (const action of ["Verify & close", "Save changes"]) {
    for (const reloadCount of [1, 3]) {
      test(`${transformedInto} ${action} rejects closing before enqueueing across ${reloadCount} reloads`, async () => {
        let stored = issue({
          transformedInto, status: transformedInto === "OFI" ? "In Progress" : "Under Testing/Revision",
          implementationDate: addCalendarDays(todayDate(), -90), verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com",
        });
        const failedAppend = deferred<void>();
        const retriedAppend = deferred<void>();
        const reloads = Array.from({ length: reloadCount }, () => deferred<void>());
        let appends = 0;
        let reads = 0;
        let closures = 0;
        await renderApp({
          addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
            const attempt = ++appends;
            if (attempt === 2) await failedAppend.promise;
            if (attempt === 3) await retriedAppend.promise;
            const saved = { ...entry, id: attempt, saveWarning: attempt === 1 ? "Your update was posted. Reload its server details." : undefined };
            stored = { ...stored, progressLog: [...stored.progressLog, saved] };
            return saved;
          },
          getIssue: async () => {
            const snapshot = { ...stored, progressLog: [...stored.progressLog] };
            await reloads[reads++].promise;
            return snapshot;
          },
          updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
            closures += 1;
            assert.equal(patch.status, "Closed");
            assert.equal(eTag, '"1"');
            stored = { ...stored, ...patch, eTag: '"2"' };
            return stored;
          },
        }, [stored]);
        await click("Issue register");
        await openIssue();
        writeProgress("Accepted observation");
        await click("Add update");
        assertSavedWarning();
        writeProgress("Rejected note preserved through reloads");
        if (action === "Save changes") change("Status", "Closed");
        const add = button("Add update");
        const reload = button("Reload saved issue");
        const close = button(action);
        assert.equal(close.disabled, false);
        await act(async () => {
          Simulate.click(add);
          for (let i = 0; i < reloadCount; i++) Simulate.click(reload);
          Simulate.click(close);
        });
        assert.equal(closures, 0);
        assert.equal(reads, 0);
        assert.equal(button(action).disabled, true);
        assert.match(container.querySelector('[role="alert"]')?.textContent || "", /pending work.*Try closing again/);
        await act(async () => { failedAppend.reject(new Error("Journal append rejected")); });
        assert.equal(closures, 0);
        assert.equal(progressInput().disabled, true);
        assert.equal(progressInput().value, "Rejected note preserved through reloads");
        assert.match(container.querySelector('p[role="alert"]')?.textContent || "", /Journal append rejected/);
        for (let i = 0; i < reloadCount; i++) {
          assert.equal(reads, i + 1);
          assert.equal(button(action).disabled, true);
          await act(async () => { button(action).click(); reloads[i].resolve(); });
          assert.equal(closures, 0);
          assert.notEqual(stored.status, "Closed");
          assert.equal(progressInput().value, "Rejected note preserved through reloads");
          assert.match(container.querySelector('p[role="alert"]')?.textContent || "", /Journal append rejected/);
        }
        assert.equal(progressInput().disabled, false);
        assert.equal(button(action).disabled, false);
        assert.deepEqual(journalText(), ["Accepted observation"]);
        await click("Add update");
        assert.equal(button(action).disabled, true);
        await act(async () => { retriedAppend.resolve(); });
        assert.equal(closures, 0);
        assert.equal(progressInput().value, "");
        await click(action);
        assert.equal(closures, 1);
        assert.equal(stored.status, "Closed");
        assert.deepEqual(journalText().sort(), ["Accepted observation", "Rejected note preserved through reloads"]);
        assert.equal(container.querySelector("textarea"), null);
        assert.equal(container.querySelector('[role="alert"]'), null);
      });
    }
  }
}

test("reload and append use the same issue queue so a stale read cannot erase an accepted entry", async () => {
  const pendingRead = deferred<void>();
  let stored = issue();
  const calls: string[] = [];
  await renderApp({
    addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
      calls.push("append");
      const id = stored.progressLog.length + 1;
      const saved = { ...entry, id, saveWarning: id === 1 ? "Your update was posted. Reload its server details." : undefined };
      stored = { ...stored, progressLog: [...stored.progressLog, saved] };
      return saved;
    },
    getIssue: async () => {
      calls.push("reload");
      const snapshot = { ...stored, progressLog: [...stored.progressLog] };
      await pendingRead.promise;
      return snapshot;
    },
  });
  await click("Issue register");
  await openIssue();
  writeProgress("First accepted observation");
  await click("Add update");
  writeProgress("Second accepted observation");
  await act(async () => { Simulate.click(button("Reload saved issue")); Simulate.click(button("Add update")); });
  assertFieldsDisabled(true);
  assert.equal(progressInput().disabled, true);
  await act(async () => { pendingRead.resolve(); });
  assert.deepEqual(calls, ["append", "reload", "append"]);
  assert.deepEqual(journalText().sort(), ["First accepted observation", "Second accepted observation"]);
  assert.equal(progressInput().value, "");
});

test("operations for another issue remain independent of a pending save", async () => {
  const pending = deferred<void>();
  const first = issue();
  const second = issue({ id: 2, qsNumber: 1002 });
  const writes: number[] = [];
  await renderApp({ updateIssue: async (id: number, patch: Partial<IIssue>) => {
    writes.push(id);
    if (id === first.id) await pending.promise;
    return { ...(id === first.id ? first : second), ...patch, eTag: '"2"' };
  } }, [first, second]);
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "First issue note");
  await click("Save changes");
  await click("Back to register");
  const secondRow = Array.from(container.querySelectorAll("button")).find(candidate => candidate.textContent?.includes("QS-1002"));
  assert.ok(secondRow);
  await act(async () => { Simulate.click(secondRow); });
  assertFieldsDisabled(false);
  assert.equal(progressInput().disabled, false);
  change("Follow up (Quality Team notes)", "Second issue note");
  await click("Save changes");
  assert.deepEqual(writes, [1, 2]);
  assert.equal(button("Save changes").disabled, true);
  await act(async () => { pending.resolve(); });
  assert.equal(field("Follow up (Quality Team notes)").value, "Second issue note");
});

test("closing another issue is allowed while an unrelated issue has pending work", async () => {
  const pending = deferred<void>();
  const first = issue();
  const second = issue({ id: 2, qsNumber: 1002, transformedInto: "OFI" });
  const writes: number[] = [];
  await renderApp({ updateIssue: async (id: number, patch: Partial<IIssue>) => {
    writes.push(id);
    if (id === first.id) await pending.promise;
    return { ...(id === first.id ? first : second), ...patch, eTag: '"2"' };
  } }, [first, second]);
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "First issue note");
  await click("Save changes");
  await click("Back to register");
  const secondRow = Array.from(container.querySelectorAll("button")).find(candidate => candidate.textContent?.includes("QS-1002"));
  assert.ok(secondRow);
  await act(async () => { Simulate.click(secondRow); });
  await click("Verify & close");
  assert.deepEqual(writes, [1, 2]);
  assert.equal(button("Re-open issue").disabled, false);
  await act(async () => { pending.resolve(); });
  assert.equal(button("Re-open issue").disabled, false);
});

for (const transformedInto of ["OFI", "NC Minor"] as const) {
  for (const action of ["Verify & close", "Save changes"]) {
    for (const fails of [false, true]) {
      test(`${transformedInto} ${action} stays locked across remount until ${fails ? "rejection and explicit retry" : "accepted closure"}`, async () => {
        let stored = issue({
          transformedInto, status: transformedInto === "OFI" ? "In Progress" : "Under Testing/Revision",
          implementationDate: addCalendarDays(todayDate(), -90), verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com",
        });
        const pending = deferred<void>();
        let writes = 0;
        let appends = 0;
        await renderApp({
          updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
            writes += 1;
            assert.equal(patch.status, "Closed");
            assert.equal(eTag, '"1"');
            if (writes === 1) await pending.promise;
            stored = { ...stored, ...patch, eTag: '"2"' };
            return stored;
          },
          addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => { appends += 1; return entry; },
        }, [stored]);
        await click("Issue register");
        await openIssue();
        if (action === "Save changes") change("Status", "Closed");
        await click(action);
        await click("Back to register");
        await openIssue();
        assertFieldsDisabled(true);
        assert.equal(progressInput().disabled, true);
        for (const label of ["Save changes", "Verify & close", "Put on hold", "Reject issue", "Add update"]) {
          assert.equal(button(label).disabled, true);
          act(() => { button(label).click(); });
        }
        assert.equal(writes, 1);
        assert.equal(appends, 0);
        assert.equal(progressInput().value, "");
        if (fails) {
          await act(async () => { pending.reject(new Error("Closure rejected after remount")); });
          assertFieldsDisabled(false);
          assert.equal(progressInput().disabled, false);
          assert.notEqual(stored.status, "Closed");
          assert.match(container.querySelector('[role="alert"]')?.textContent || "", /Closure rejected after remount/);
          if (action === "Save changes") change("Status", "Closed");
          await click(action);
        } else {
          await act(async () => { pending.resolve(); });
        }
        assert.equal(writes, fails ? 2 : 1);
        assert.equal(appends, 0);
        assert.equal(container.querySelector("textarea"), null);
        assert.equal(container.querySelector('[role="alert"]'), null);
        assert.equal(button("Re-open issue").disabled, false);
        assert.equal(stored.status, "Closed");
      });
    }
  }
}

for (const profile of ["qm", "owner"] as const) {
  const identity = { profile, userDisplayName: "Owner", userEmail: "owner@example.com" };
  const back = profile === "owner" ? "Back to my tasks" : "Back to register";
  const label = profile === "owner" ? "Status" : "Follow up (Quality Team notes)";
  const submit = profile === "owner" ? "Save" : "Save changes";

  test(`${profile} remounted during a non-closing save resumes with the accepted fields and ETag`, async () => {
    let stored = issue({ transformedInto: "OFI", status: "Created", followUp: "Original note" });
    const pending = deferred<void>();
    const writes: { patch: Partial<IIssue>; eTag: string }[] = [];
    await renderApp({ updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
      writes.push({ patch, eTag });
      if (writes.length === 1) await pending.promise;
      stored = { ...stored, ...patch, eTag: `"${writes.length + 1}"` };
      return stored;
    } }, [stored], identity);
    await click("Issue register");
    await openIssue();
    change(label, profile === "owner" ? "In Progress" : "Accepted note");
    await click(submit);
    await click(back);
    await openIssue();
    assertFieldsDisabled(true);
    assert.equal(progressInput().disabled, true);
    assert.equal(field(label).value, profile === "owner" ? "Created" : "Original note");
    await act(async () => { pending.resolve(); });
    assertFieldsDisabled(false);
    assert.equal(progressInput().disabled, false);
    assert.equal(field(label).value, profile === "owner" ? "In Progress" : "Accepted note");
    assert.equal(button(submit).disabled, true);
    change(label, profile === "owner" ? "Created" : "Next note");
    await click(submit);
    assert.deepEqual(writes.map(write => write.eTag), ['"1"', '"2"']);
    if (profile === "qm") assert.deepEqual(writes[1].patch, { followUp: "Next note" });
    else assert.equal(writes[1].patch.status, "Created");
    assert.equal(container.querySelector('[role="alert"]'), null);
  });

  test(`${profile} remounted during an append keeps all issue controls locked`, async () => {
    const pending = deferred<IProgressLogEntry>();
    const entry = { id: 31, text: "Accepted before navigation", author: "Owner", ts: new Date().toISOString() };
    let appends = 0;
    await renderApp({ addProgressLogEntry: () => { appends += 1; return pending.promise; } }, [issue()], identity);
    await click("Issue register");
    await openIssue();
    writeProgress(entry.text);
    await click("Add update");
    await click(back);
    await openIssue();
    assertFieldsDisabled(true);
    assert.equal(progressInput().disabled, true);
    assert.equal(button("Add update").disabled, true);
    await act(async () => { pending.resolve(entry); });
    assertFieldsDisabled(false);
    assert.equal(progressInput().disabled, false);
    assert.deepEqual(journalText(), [entry.text]);
    assert.equal(appends, 1);
  });

  test(`${profile} retains genuine edits and their original ETag when reload brings a newer snapshot`, async () => {
    const original = issue({ transformedInto: "OFI", status: "Created", followUp: "Original note" });
    const entry = { id: 32, text: "Accepted observation", author: "Owner", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
    const latest = { ...original, followUp: "Another manager's note", progressLog: [entry], eTag: '"2"' };
    const pending = deferred<IIssue>();
    const writes: { patch: Partial<IIssue>; eTag: string }[] = [];
    let reads = 0;
    await renderApp({
      addProgressLogEntry: async () => entry,
      getIssue: async () => { reads += 1; return reads === 1 ? pending.promise : latest; },
      updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => { writes.push({ patch, eTag }); throw new IssueConflictError(1, latest); },
    }, [original], identity);
    await click("Issue register");
    await openIssue();
    writeProgress(entry.text);
    await click("Add update");
    change(label, profile === "owner" ? "In Progress" : "Genuine unsaved note");
    writeProgress("Unsubmitted journal draft");
    await click("Reload saved issue");
    assertFieldsDisabled(true);
    assert.equal(progressInput().disabled, true);
    await act(async () => { pending.resolve(latest); });
    assertFieldsDisabled(false);
    assert.equal(field(label).value, profile === "owner" ? "In Progress" : "Genuine unsaved note");
    assert.equal(progressInput().value, "Unsubmitted journal draft");
    assert.equal(progressInput().disabled, false);
    await click(submit);
    assert.equal(writes.length, 1);
    assert.equal(writes[0].eTag, '"1"');
    if (profile === "qm") assert.deepEqual(writes[0].patch, { followUp: "Genuine unsaved note" });
    else assert.equal(writes[0].patch.status, "In Progress");
    assert.equal(field(label).value, profile === "owner" ? "In Progress" : "Genuine unsaved note");
    assert.match(container.textContent || "", /Your draft has been kept/);
    await click("Reload latest and discard draft");
    assert.equal(field(label).value, profile === "owner" ? "Created" : "Another manager's note");
    assert.equal(progressInput().value, "Unsubmitted journal draft");
  });
}

test("triage remounted during an accepted write requires recovery before another submission", async () => {
  const pending = deferred<IIssue>();
  const original = issue({ triaged: false, status: undefined });
  let writes = 0;
  await renderApp({
    updateIssue: () => { writes += 1; return pending.promise; },
    getIssue: async () => ({ ...original, triaged: true, status: "Created", eTag: '"2"' }),
  }, [original]);
  await click("Triage queue1");
  await openIssue();
  change("Task owner (gets reminders)", "Owner");
  change("Task owner Microsoft 365 email", "owner@example.com");
  await click("Create issue");
  await click("Back to triage queue");
  await openIssue();
  assertFieldsDisabled(true);
  assert.equal(button("Create issue").disabled, true);
  assert.equal(button("Reject (no action)").disabled, true);
  await act(async () => { pending.reject(new IssueRefreshError(original.id)); });
  assertSavedWarning();
  assert.equal(button("Create issue").disabled, true);
  await click("Reload latest and return to queue");
  await click("Issue register");
  await openIssue();
  assertFieldsDisabled(false);
  assert.equal(field("Status").value, "Created");
  assert.equal(writes, 1);
});

for (const operation of ["triage", "reload"] as const) {
  test(`${operation} completion for issue A cannot navigate away from issue B's drafts`, async () => {
    const first = issue({ triaged: false, status: undefined, eTag: operation === "reload" ? undefined : '"1"' });
    const second = issue({ id: 2, qsNumber: 1002 });
    const pending = deferred<void>();
    const calls: string[] = [];
    await renderApp({
      updateIssue: async (id: number, patch: Partial<IIssue>) => {
        assert.equal(id, first.id);
        calls.push("triage");
        await pending.promise;
        return { ...first, ...patch, eTag: '"2"' };
      },
      getIssue: async (id: number) => {
        assert.equal(id, first.id);
        calls.push("reload");
        await pending.promise;
        return { ...first, eTag: '"2"' };
      },
    }, [first, second]);
    await click("Triage queue1");
    await openIssue(1001);
    if (operation === "triage") {
      change("Task owner (gets reminders)", "Owner");
      change("Task owner Microsoft 365 email", "owner@example.com");
      await click("Create issue");
    } else await click("Reload latest and return to queue");
    await click("Back to triage queue");
    await click("Issue register");
    await openIssue(1002);
    assertFieldsDisabled(false);
    change("Follow up (Quality Team notes)", "Issue B's unsaved assessment");
    writeProgress("Issue B's unposted evidence");
    await act(async () => { pending.resolve(); });
    assert.equal(field("Follow up (Quality Team notes)").value, "Issue B's unsaved assessment");
    assert.equal(progressInput().value, "Issue B's unposted evidence");
    assert.equal(button("Save changes").disabled, false);
    assert.equal(button("Add update").disabled, false);
    assert.deepEqual(calls, [operation]);
  });
}

function recoveryPanel(): HTMLElement | null {
  return container.querySelector('section[aria-label="Unsaved draft recovery"]');
}

function recoveredFields(panel = recoveryPanel()): Record<string, string> {
  assert.ok(panel, "Unsaved draft recovery should be available");
  return Object.fromEntries(Array.from(panel.querySelectorAll("dt")).map(label => [label.textContent, label.nextElementSibling?.textContent]));
}

function recoveredCopies(): Record<string, string>[] {
  return Array.from(container.querySelectorAll<HTMLElement>('section[aria-label="Unsaved draft recovery"]')).map(panel => recoveredFields(panel));
}

async function discardRecoveredCopy(index: number): Promise<void> {
  const panel = container.querySelectorAll('section[aria-label="Unsaved draft recovery"]')[index];
  assert.ok(panel);
  const discard = panel.querySelector("button");
  assert.ok(discard);
  assert.equal(discard.textContent, "Discard recovered draft");
  await act(async () => { Simulate.click(discard); });
}

for (const [profile, transition] of [["qm", "Closed"], ["owner", "Closed"], ["owner", "reassigned"]] as const) {
  test(`${profile} retains selectable detail and progress recovery after an external ${transition} refresh`, async () => {
    const original = issue({ status: "Created", followUp: "Original note", eTag: '"private-version-1"' });
    const second = issue({ id: 2, qsNumber: 1002 });
    const accepted = { id: 40, text: "Already accepted evidence", author: "Owner", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
    const latest = {
      ...original, eTag: '"private-version-2"', progressLog: [accepted],
      ...(transition === "Closed" ? { status: "Closed" as const } : { taskOwnerEmail: "replacement@example.com" }),
    };
    const pending = deferred<IIssue>();
    let writes = 0;
    let appends = 0;
    await renderApp({
      addProgressLogEntry: async () => { appends += 1; return accepted; },
      getIssue: () => pending.promise,
      updateIssue: async () => { writes += 1; throw new Error("No write should occur"); },
    }, [original, second], { profile, userDisplayName: "Owner", userEmail: "owner@example.com" });
    await click("Issue register");
    await openIssue(1001);
    writeProgress(accepted.text);
    await click("Add update");
    const date = addCalendarDays(todayDate(), -3);
    if (profile === "qm") {
      change("Follow up (Quality Team notes)", "");
      change("Root cause", "Unsaved assessment\nSecond line");
      change("Task owner Microsoft 365 email", "draft@example.com");
    } else {
      change("Status", "In Progress");
      change("Implementation date", date);
    }
    writeProgress("Unposted progress\nKeep this evidence");
    await click("Reload saved issue");
    assertFieldsDisabled(true);
    await act(async () => { pending.resolve(latest); });
    const expected = {
      ...(profile === "qm" ? { "Quality Team notes": "(Cleared)", "Root cause": "Unsaved assessment\nSecond line", "Task owner Microsoft 365 email": "draft@example.com" } : { Status: "In Progress", "Implementation date": date }),
      "Unposted progress note": "Unposted progress\nKeep this evidence",
    };
    assert.deepEqual(recoveredFields(), expected);
    assert.equal(container.querySelectorAll("label input,label select,label textarea,textarea[placeholder^='What did you do']").length, 0);
    assert.equal(Array.from(container.querySelectorAll("button")).some(element => ["Save changes", "Save progress", "Add update", "Verify & close"].includes(element.textContent || "")), false);
    if (profile === "owner") assert.equal(Array.from(container.querySelectorAll("button")).some(element => element.textContent === "Re-open issue"), false);
    assert.doesNotMatch(recoveryPanel()?.textContent || "", /private-version|eTag|Already accepted evidence/);
    const content = recoveryPanel()?.querySelector("dl");
    assert.ok(content);
    assert.equal(window.getComputedStyle(content).userSelect, "text");
    await click("Back to register");
    await openIssue(1002);
    assert.equal(recoveryPanel(), null);
    assertFieldsDisabled(false);
    await click(profile === "owner" ? "Back to my tasks" : "Back to register");
    await openIssue(1001);
    assert.deepEqual(recoveredFields(), expected);
    await click("Discard recovered draft");
    assert.equal(recoveryPanel(), null);
    await click("Back to register");
    await openIssue(1001);
    assert.equal(recoveryPanel(), null);
    assert.equal(writes, 0);
    assert.equal(appends, 1);
  });
}

for (const transformedInto of ["OFI", "NC Minor"] as const) {
  for (const action of ["Verify & close", "Save changes"]) {
    for (const readbackFails of [false, true]) {
      test(`${transformedInto} ${action} acceptance ${readbackFails ? "with failed readback" : "with readback"} immediately closes the App view and recovers only unsubmitted progress`, async () => {
        let stored = issue({
          transformedInto, status: transformedInto === "OFI" ? "In Progress" : "Under Testing/Revision",
          implementationDate: addCalendarDays(todayDate(), -90), verifiedBy: "Verifier", verifiedByEmail: "verifier@example.com",
        });
        const writes: { patch: Partial<IIssue>; eTag: string }[] = [];
        let reads = 0;
        let appends = 0;
        await renderApp({
          updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
            writes.push({ patch, eTag });
            stored = { ...stored, ...patch, eTag: `"${writes.length + 1}"` };
            if (readbackFails && writes.length === 1) throw new IssueRefreshError(stored.id);
            return stored;
          },
          getIssue: async () => {
            reads += 1;
            if (reads === 1) throw new Error("Read still unavailable");
            return stored;
          },
          addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => { appends += 1; return { ...entry, id: appends }; },
        }, [stored]);
        await click("Issue register");
        await openIssue();
        change("Follow up (Quality Team notes)", "Submitted assessment");
        writeProgress("Unsubmitted progress to recover");
        if (action === "Save changes") change("Status", "Closed");
        await click(action);
        assert.equal(stored.status, "Closed");
        assert.match(container.textContent || "", /Closed · read-only/);
        assert.equal(container.querySelector("textarea"), null);
        assert.deepEqual(recoveredFields(), { "Unposted progress note": "Unsubmitted progress to recover" });
        assert.doesNotMatch(recoveryPanel()?.textContent || "", /Submitted assessment/);
        assert.match(container.textContent || "", /Submitted assessment/);
        assert.ok(Array.from(container.querySelectorAll("dt")).some(element => element.textContent === "Closed on"));
        assert.equal(writes.length, 1);
        assert.equal(writes[0].eTag, '"1"');
        assert.equal(appends, 0);
        if (readbackFails) {
          assertSavedWarning();
          assert.equal(button("Re-open issue").disabled, true);
          act(() => { button("Re-open issue").click(); });
          await click("Back to register");
          await openIssue();
          assert.match(container.textContent || "", /Closed · read-only/);
          assert.equal(container.querySelector("textarea"), null);
          assert.equal(button("Re-open issue").disabled, true);
          assert.deepEqual(recoveredFields(), { "Unposted progress note": "Unsubmitted progress to recover" });
          await click("Reload saved issue");
          assertSavedWarning();
          assert.match(container.querySelector('[role="status"]')?.textContent || "", /Read still unavailable/);
          assert.equal(container.querySelector("textarea"), null);
          assert.equal(button("Re-open issue").disabled, true);
          assert.equal(writes.length, 1);
          assert.equal(appends, 0);
          await click("Reload saved issue");
          assert.equal(reads, 2);
          assert.equal(container.querySelector('[role="status"]'), null);
          assert.equal(container.querySelector("textarea"), null);
          assert.equal(button("Re-open issue").disabled, false);
          assert.deepEqual(recoveredFields(), { "Unposted progress note": "Unsubmitted progress to recover" });
          await click("Re-open issue");
          assert.equal(writes.length, 2);
          assert.equal(writes[1].eTag, '"2"');
          assert.equal(writes[1].patch.status, "In Progress");
          assert.equal(appends, 1);
          assert.equal(field("Status").value, "In Progress");
          assert.equal(progressInput().disabled, false);
          assert.equal(progressInput().value, "");
          assert.deepEqual(recoveredFields(), { "Unposted progress note": "Unsubmitted progress to recover" });
        }
      });
    }
  }
}

for (const action of ["Verify & close", "Save changes"]) {
  test(`${action} retains issue-specific reload and recovery after another issue replaces its warning`, async () => {
    let first = issue({ transformedInto: "OFI" });
    let second = issue({ id: 2, qsNumber: 1002, transformedInto: "OFI" });
    const reloads = [deferred<IIssue>(), deferred<IIssue>()];
    const readIds: number[] = [];
    let firstReads = 0;
    let writes = 0;
    let appends = 0;
    await renderApp({
      updateIssue: async (id: number, patch: Partial<IIssue>, eTag: string) => {
        assert.equal(id, first.id);
        assert.equal(eTag, '"1"');
        assert.equal(patch.status, "Closed");
        writes += 1;
        first = { ...first, ...patch, eTag: '"2"' };
        throw new IssueRefreshError(id);
      },
      addProgressLogEntry: async (id: number, entry: IProgressLogEntry) => {
        assert.equal(id, second.id);
        appends += 1;
        const saved = { ...entry, id: appends, saveWarning: `Issue B journal entry ${appends} was saved; reload its details.` };
        second = { ...second, progressLog: [...second.progressLog, saved] };
        return saved;
      },
      getIssue: (id: number) => {
        readIds.push(id);
        if (id === first.id) return reloads[firstReads++].promise;
        assert.equal(id, second.id);
        return Promise.resolve(second);
      },
    }, [first, second]);
    await click("Issue register");
    await openIssue(1001);
    change("Follow up (Quality Team notes)", "Accepted A assessment");
    writeProgress("Keep A's unposted evidence");
    if (action === "Save changes") change("Status", "Closed");
    await click(action);
    assertSavedWarning();
    const recovered = [{ "Unposted progress note": "Keep A's unposted evidence" }];
    assert.deepEqual(recoveredCopies(), recovered);
    assert.equal(button("Re-open issue").disabled, true);
    await click("Back to register");
    await openIssue(1002);
    writeProgress("B's first observation");
    await click("Add update");
    assert.match(container.querySelector('[role="status"]')?.textContent || "", /Issue B journal entry 1 was saved/);
    await click("Reload saved issue");
    assert.equal(container.querySelector('[role="status"]'), null);
    assert.deepEqual(readIds, [second.id]);
    await click("Back to register");
    await openIssue(1001);
    assert.deepEqual(recoveredCopies(), recovered);
    assert.equal(container.querySelector('[role="status"]'), null);
    assert.equal(container.querySelector("textarea"), null);
    assert.equal(button("Re-open issue").disabled, true);
    await click("Reload this issue");
    assert.equal(button("Reload this issue").disabled, true);
    assert.equal(button("Re-open issue").disabled, true);
    act(() => { button("Reload this issue").click(); button("Re-open issue").click(); });
    assert.equal(firstReads, 1);
    assert.deepEqual(recoveredCopies(), recovered);
    await act(async () => { reloads[0].reject(new Error("Issue A reload unavailable")); });
    assert.match(container.querySelector('[role="alert"]')?.textContent || "", /Could not reload this issue: Issue A reload unavailable/);
    assert.doesNotMatch(container.textContent || "", /failed and was not saved/);
    assert.match(container.textContent || "", /Closed · read-only/);
    assert.equal(container.querySelector("textarea"), null);
    assert.equal(button("Re-open issue").disabled, true);
    assert.equal(button("Reload this issue").disabled, false);
    assert.deepEqual(recoveredCopies(), recovered);
    await click("Back to register");
    await openIssue(1002);
    writeProgress("B's second observation");
    await click("Add update");
    await click("Back to register");
    await openIssue(1001);
    assert.match(container.querySelector('[role="status"]')?.textContent || "", /Issue B journal entry 2 was saved/);
    assert.deepEqual(recoveredCopies(), recovered);
    await click("Reload this issue");
    assert.equal(button("Reload this issue").disabled, true);
    assert.equal(button("Re-open issue").disabled, true);
    assert.deepEqual(recoveredCopies(), recovered);
    await act(async () => { reloads[1].resolve(first); });
    assert.deepEqual(readIds, [second.id, first.id, first.id]);
    assert.equal(container.querySelector('[role="alert"]'), null);
    assert.match(container.querySelector('[role="status"]')?.textContent || "", /Issue B journal entry 2 was saved/);
    assert.equal(button("Re-open issue").disabled, false);
    assert.equal(Array.from(container.querySelectorAll("button")).some(element => element.textContent === "Reload this issue"), false);
    assert.equal(container.querySelector("textarea"), null);
    assert.deepEqual(recoveredCopies(), recovered);
    await click("Back to register");
    await openIssue(1001);
    assert.deepEqual(recoveredCopies(), recovered);
    assert.equal(button("Re-open issue").disabled, false);
    assert.equal(writes, 1);
    assert.equal(appends, 2);
  });
}

for (const [profile, eTag] of [["owner", undefined], ["reader", "*"]] as const) {
  test(`${profile} can reload a read-only issue with an unusable version without gaining edit controls`, async () => {
    const original = issue({ status: "Closed", eTag });
    let reads = 0;
    const services = {
      getIssue: async (id: number) => { assert.equal(id, original.id); reads += 1; return { ...original, eTag: '"2"' }; },
    };
    await renderApp(services, [original]);
    await click("Issue register");
    await openIssue();
    assert.equal(button("Re-open issue").disabled, true);
    await renderApp(services, [original], { profile, userDisplayName: "Owner", userEmail: "owner@example.com" });
    assert.equal(container.querySelector('[role="status"]'), null);
    assert.equal(button("Reload this issue").disabled, false);
    await click("Reload this issue");
    assert.equal(reads, 1);
    assert.equal(container.querySelector("textarea"), null);
    assert.equal(Array.from(container.querySelectorAll("button")).some(element => ["Reload this issue", "Re-open issue", "Save changes", "Add update"].includes(element.textContent || "")), false);
  });
}

for (const action of ["Verify & close", "Save changes"]) {
  for (const readbackFails of [false, true]) {
    test(`${action} ${readbackFails ? "accepted readback failure" : "accepted closure"} across remount blocks an already queued append`, async () => {
      let stored = issue({ transformedInto: "OFI" });
      const readback = deferred<void>();
      let writes = 0;
      let appends = 0;
      await renderApp({
        updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
          writes += 1;
          assert.equal(eTag, '"1"');
          stored = { ...stored, ...patch, eTag: '"2"' };
          await readback.promise;
          return stored;
        },
        addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => { appends += 1; return entry; },
        getIssue: async () => stored,
      }, [stored]);
      await click("Issue register");
      await openIssue();
      change("Follow up (Quality Team notes)", "Accepted before remount");
      writeProgress("Queued progress must remain recoverable");
      if (action === "Save changes") change("Status", "Closed");
      await act(async () => { Simulate.click(button(action)); Simulate.click(button("Add update")); });
      assert.equal(writes, 1);
      assert.equal(appends, 0);
      await click("Back to register");
      await openIssue();
      assertFieldsDisabled(true);
      assert.equal(progressInput().disabled, true);
      await act(async () => {
        if (readbackFails) readback.reject(new IssueRefreshError(stored.id));
        else readback.resolve();
      });
      assert.equal(appends, 0);
      assert.equal(writes, 1);
      assert.match(container.textContent || "", /Closed · read-only/);
      assert.equal(container.querySelector("textarea"), null);
      assert.equal(container.querySelector('[role="alert"]'), null);
      assert.deepEqual(journalText(), []);
      assert.deepEqual(recoveredFields(), { "Unposted progress note": "Queued progress must remain recoverable" });
      if (readbackFails) {
        assertSavedWarning();
        assert.equal(button("Re-open issue").disabled, true);
        await click("Reload saved issue");
      }
      await click("Back to register");
      await openIssue();
      assert.equal(button("Re-open issue").disabled, false);
      assert.equal(container.querySelector("textarea"), null);
      assert.deepEqual(recoveredFields(), { "Unposted progress note": "Queued progress must remain recoverable" });
      assert.equal(appends, 0);
    });
  }
}

test("accepted append clears only its submitted text while a detail draft remains recoverable", async () => {
  const original = issue();
  const entry = { id: 41, text: "Accepted progress", author: "Quality Manager", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
  await renderApp({
    addProgressLogEntry: async (_id: number, submitted: IProgressLogEntry) => { assert.equal(submitted.text, entry.text); return entry; },
    getIssue: async () => ({ ...original, status: "Closed", eTag: '"2"', progressLog: [entry] }),
  }, [original]);
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Unsubmitted assessment");
  writeProgress("  Accepted progress  ");
  await click("Add update");
  await click("Reload saved issue");
  assert.deepEqual(recoveredFields(), { "Quality Team notes": "Unsubmitted assessment" });
  assert.deepEqual(journalText(), [entry.text]);
});

test("accepted detail and progress submissions do not leave false recovery content", async () => {
  let stored = issue();
  const entry = { id: 42, text: "Accepted progress", author: "Quality Manager", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
  await renderApp({
    updateIssue: async (_id: number, patch: Partial<IIssue>) => { stored = { ...stored, ...patch, eTag: '"2"' }; return stored; },
    addProgressLogEntry: async () => { stored = { ...stored, progressLog: [entry] }; return entry; },
    getIssue: async () => ({ ...stored, status: "Closed", eTag: '"3"' }),
  }, [stored]);
  await click("Issue register");
  await openIssue();
  change("Follow up (Quality Team notes)", "Accepted assessment");
  writeProgress(entry.text);
  await click("Save changes");
  assert.equal(progressInput().value, entry.text);
  await click("Add update");
  await click("Reload saved issue");
  assert.equal(recoveryPanel(), null);
  assert.equal(container.querySelector("textarea"), null);
  assert.deepEqual(journalText(), [entry.text]);
});

test("a hold dialog draft survives a refresh that closes its issue", async () => {
  const original = issue();
  const entry = { id: 43, text: "Accepted progress", author: "Quality Manager", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
  await renderApp({
    addProgressLogEntry: async () => entry,
    getIssue: async () => ({ ...original, status: "Closed", eTag: '"2"', progressLog: [entry] }),
  }, [original]);
  await click("Issue register");
  await openIssue();
  writeProgress(entry.text);
  await click("Add update");
  await click("Put on hold");
  const until = addCalendarDays(todayDate(), 7);
  change("Reason for hold *", "Waiting for replacement equipment");
  change("Resume work on *", until);
  await click("Reload saved issue");
  assert.deepEqual(recoveredFields(), { "Reason for hold": "Waiting for replacement equipment", "Resume work on": until });
  assert.equal(container.querySelector("textarea"), null);
});

test("rejected detail and progress submissions remain recoverable after external closure", async () => {
  const original = issue();
  const entry = { id: 44, text: "Accepted observation", author: "Quality Manager", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
  let appends = 0;
  await renderApp({
    addProgressLogEntry: async () => { if (++appends > 1) throw new Error("Append rejected"); return entry; },
    updateIssue: async () => { throw new Error("Detail rejected"); },
    getIssue: async () => ({ ...original, status: "Closed", eTag: '"2"', progressLog: [entry] }),
  }, [original]);
  await click("Issue register");
  await openIssue();
  writeProgress(entry.text);
  await click("Add update");
  change("Follow up (Quality Team notes)", "Rejected assessment");
  await click("Save changes");
  writeProgress("Rejected progress");
  await click("Add update");
  await click("Reload saved issue");
  assert.deepEqual(recoveredFields(), { "Quality Team notes": "Rejected assessment", "Unposted progress note": "Rejected progress" });
  assert.equal(container.querySelector("textarea"), null);
});

test("accepted hold fields do not leave a recovery draft", async () => {
  let stored = issue({ transformedInto: "OFI" });
  const entry = { id: 45, text: "Accepted observation", author: "Quality Manager", ts: new Date().toISOString(), saveWarning: "Your update was posted. Reload its server details." };
  await renderApp({
    updateIssue: async (_id: number, patch: Partial<IIssue>) => { stored = { ...stored, ...patch, eTag: '"2"' }; return stored; },
    addProgressLogEntry: async () => entry,
    getIssue: async () => ({ ...stored, status: "Closed", eTag: '"3"', progressLog: [entry] }),
  }, [stored]);
  await click("Issue register");
  await openIssue();
  await click("Put on hold");
  change("Reason for hold *", "  Awaiting parts  ");
  change("Resume work on *", addCalendarDays(todayDate(), 7));
  const confirms = Array.from(container.querySelectorAll("button")).filter(candidate => candidate.textContent === "Put on hold");
  await act(async () => { Simulate.click(confirms[confirms.length - 1]); });
  assert.equal(stored.holdReason, "Awaiting parts");
  writeProgress(entry.text);
  await click("Add update");
  await click("Reload saved issue");
  assert.equal(recoveryPanel(), null);
});

test("reopened recovery stays visible through initialization, audit writes and matching fresh submissions", async () => {
  let stored = issue({ transformedInto: "OFI", status: "Created" });
  const patches: { patch: Partial<IIssue>; eTag: string }[] = [];
  const entries: IProgressLogEntry[] = [];
  const auditText = "Issue re-opened — corrective-action cycle restarted. Previous close: —.";
  let failNextSave = false;
  await renderApp({
    getIssue: async () => stored,
    updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
      patches.push({ patch, eTag });
      if (failNextSave) { failNextSave = false; throw new IssueConflictError(stored.id, stored); }
      stored = { ...stored, ...patch, eTag: `"${patches.length + 2}"` };
      return stored;
    },
    addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
      const accepted = { ...entry, id: entries.length + 1, saveWarning: "Your update was posted. Reload its server details." };
      entries.push(accepted);
      stored = { ...stored, progressLog: [...entries] };
      return accepted;
    },
  }, [stored]);
  await click("Issue register");
  await openIssue();
  writeProgress("Accepted observation");
  await click("Add update");
  change("Status", "In Progress");
  change("Follow up (Quality Team notes)", "Recovered assessment");
  writeProgress(auditText);
  stored = { ...stored, status: "Closed", eTag: '"2"' };
  await click("Reload saved issue");
  const recovered = { Status: "In Progress", "Quality Team notes": "Recovered assessment", "Unposted progress note": auditText };
  assert.deepEqual(recoveredCopies(), [recovered]);
  assert.equal(patches.length, 0);
  await click("Re-open issue");
  assert.deepEqual(recoveredCopies(), [recovered]);
  assert.equal(field("Follow up (Quality Team notes)").value, "");
  assert.equal(progressInput().value, "");
  assert.equal(entries[1].text, auditText);
  assert.equal(patches[0].patch.followUp, undefined);
  assert.equal(patches[0].eTag, '"2"');
  change("Follow up (Quality Team notes)", "Recovered assessment");
  writeProgress(auditText);
  assert.deepEqual(recoveredCopies(), [recovered]);
  await click("Save changes");
  assert.equal(patches[1].eTag, '"3"');
  assert.equal(progressInput().value, auditText);
  assert.deepEqual(recoveredCopies(), [recovered]);
  await click("Add update");
  assert.equal(progressInput().value, "");
  assert.deepEqual(recoveredCopies(), [recovered]);
  failNextSave = true;
  change("Follow up (Quality Team notes)", "A conflicting fresh assessment");
  writeProgress("Keep the live progress draft");
  await click("Save changes");
  await click("Reload latest and discard draft");
  assert.equal(field("Follow up (Quality Team notes)").value, "Recovered assessment");
  assert.equal(progressInput().value, "Keep the live progress draft");
  assert.deepEqual(recoveredCopies(), [recovered]);
  const content = recoveryPanel()?.querySelector("dl");
  assert.ok(content);
  assert.equal(window.getComputedStyle(content).userSelect, "text");
  change("Follow up (Quality Team notes)", "Keep the live detail draft");
  await discardRecoveredCopy(0);
  assert.equal(recoveryPanel(), null);
  assert.equal(field("Follow up (Quality Team notes)").value, "Keep the live detail draft");
  assert.equal(progressInput().value, "Keep the live progress draft");
  assert.equal(button("Save changes").disabled, false);
  assert.equal(button("Add update").disabled, false);
  assert.equal(patches.length, 3);
  assert.equal(entries.length, 3);
});

for (const identical of [false, true]) {
  test(`repeated close and reopen retains ${identical ? "identical" : "distinct"} recovery copies and discards only the selected copy`, async () => {
    let stored = issue({ transformedInto: "OFI" });
    let writes = 0;
    let appends = 0;
    await renderApp({
      getIssue: async () => stored,
      updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
        assert.equal(eTag, stored.eTag);
        writes += 1;
        stored = { ...stored, ...patch, eTag: `"write-${writes}"` };
        return stored;
      },
      addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
        const accepted = { ...entry, id: ++appends, saveWarning: "Your update was posted. Reload its server details." };
        stored = { ...stored, progressLog: [...stored.progressLog, accepted] };
        return accepted;
      },
    }, [stored]);
    await click("Issue register");
    await openIssue();
    writeProgress("Accepted observation");
    await click("Add update");
    const copies: Record<string, string>[] = [];
    for (let round = 0; round < 2; round++) {
      const suffix = identical ? "same" : String(round + 1);
      change("Follow up (Quality Team notes)", `Assessment ${suffix}`);
      writeProgress(`Evidence ${suffix}`);
      stored = { ...stored, status: "Closed", eTag: `"close-${round}"` };
      await click("Reload saved issue");
      copies.push({ "Quality Team notes": `Assessment ${suffix}`, "Unposted progress note": `Evidence ${suffix}` });
      assert.deepEqual(recoveredCopies(), copies);
      await click("Back to register");
      await openIssue();
      assert.deepEqual(recoveredCopies(), copies);
      await click("Re-open issue");
      assert.deepEqual(recoveredCopies(), copies);
      assert.equal(field("Follow up (Quality Team notes)").value, "");
      assert.equal(progressInput().value, "");
    }
    assert.equal(writes, 2);
    assert.equal(appends, 3);
    change("Follow up (Quality Team notes)", "Fresh live assessment");
    writeProgress("Fresh live evidence");
    await discardRecoveredCopy(0);
    assert.deepEqual(recoveredCopies(), [copies[1]]);
    assert.equal(field("Follow up (Quality Team notes)").value, "Fresh live assessment");
    assert.equal(progressInput().value, "Fresh live evidence");
    await click("Save changes");
    await click("Add update");
    assert.deepEqual(recoveredCopies(), [copies[1]]);
    stored = { ...stored, status: "Closed", eTag: '"final-close"' };
    await click("Reload saved issue");
    assert.deepEqual(recoveredCopies(), [copies[1]]);
    await discardRecoveredCopy(0);
    assert.equal(recoveryPanel(), null);
    assert.equal(writes, 3);
    assert.equal(appends, 4);
  });
}

test("owner recovery survives reassignment away and back, matching fresh writes, and another reassignment", async () => {
  let stored = issue({ status: "Created" });
  const identity = { profile: "owner", userDisplayName: "Owner", userEmail: "owner@example.com" };
  let writes = 0;
  let appends = 0;
  const services = {
    getIssue: async () => stored,
    updateIssue: async (_id: number, patch: Partial<IIssue>, eTag: string) => {
      assert.equal(eTag, stored.eTag);
      writes += 1;
      stored = { ...stored, ...patch, eTag: `"write-${writes}"` };
      return stored;
    },
    addProgressLogEntry: async (_id: number, entry: IProgressLogEntry) => {
      const accepted = { ...entry, id: ++appends, saveWarning: "Your update was posted. Reload its server details." };
      stored = { ...stored, progressLog: [...stored.progressLog, accepted] };
      return accepted;
    },
  };
  await renderApp(services, [stored], identity);
  await click("Issue register");
  await openIssue();
  writeProgress("Accepted observation");
  await click("Add update");
  const date = addCalendarDays(todayDate(), -3);
  change("Status", "In Progress");
  change("Implementation date", date);
  writeProgress("Owner evidence");
  stored = { ...stored, taskOwnerEmail: "replacement@example.com", eTag: '"away-1"' };
  await click("Reload saved issue");
  const first = { Status: "In Progress", "Implementation date": date, "Unposted progress note": "Owner evidence" };
  assert.deepEqual(recoveredCopies(), [first]);
  assert.equal(container.querySelector("textarea"), null);
  stored = { ...stored, taskOwnerEmail: "owner@example.com", eTag: '"back-1"' };
  await renderApp(services, [stored], identity);
  assert.deepEqual(recoveredCopies(), [first]);
  assert.equal(field("Status").value, "Created");
  assert.equal(progressInput().value, "");
  change("Status", "In Progress");
  change("Implementation date", date);
  writeProgress("Owner evidence");
  await click("Save progress");
  await click("Add update");
  assert.deepEqual(recoveredCopies(), [first]);
  change("Status", "Created");
  writeProgress("Second owner draft");
  stored = { ...stored, taskOwnerEmail: "replacement@example.com", eTag: '"away-2"' };
  await click("Reload saved issue");
  const second = { Status: "Created", "Unposted progress note": "Second owner draft" };
  assert.deepEqual(recoveredCopies(), [first, second]);
  stored = { ...stored, taskOwnerEmail: "owner@example.com", eTag: '"back-2"' };
  await renderApp(services, [stored], identity);
  assert.deepEqual(recoveredCopies(), [first, second]);
  assert.equal(field("Status").value, "In Progress");
  const nextDate = addCalendarDays(todayDate(), -1);
  change("Implementation date", nextDate);
  writeProgress("Fresh owner draft");
  await discardRecoveredCopy(1);
  assert.deepEqual(recoveredCopies(), [first]);
  assert.equal(field("Implementation date").value, nextDate);
  assert.equal(progressInput().value, "Fresh owner draft");
  await discardRecoveredCopy(0);
  assert.equal(recoveryPanel(), null);
  await click("Save progress");
  await click("Add update");
  assert.equal(recoveryPanel(), null);
  assert.equal(writes, 2);
  assert.equal(appends, 3);
});
