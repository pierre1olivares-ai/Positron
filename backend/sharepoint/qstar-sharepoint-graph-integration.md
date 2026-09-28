# Q-Star Issue Manager — SharePoint & Microsoft Graph integration guide

This document is the bridge between the standalone test build (`qstar-issue-manager.jsx`) and the production deployment in your repository. It covers three things:

1. The **field mapping** between the app's data model and the SharePoint list that backs the register.
2. The **Microsoft Graph data layer** that replaces the local `window.storage` calls.
3. The **Power Automate reminder flow** that reproduces the app's simulated reminder engine.

Everything here mirrors the logic already implemented in the app, so the production version behaves the same way the test build does.

---

## 1. Data flow overview

The system runs in two phases, exactly as modelled in the app.

**Phase 1 — Intake (anyone).** An employee submits the Q-Star **Microsoft Form**. A Power Automate flow (triggered by *When a new response is submitted*) creates a SharePoint list item with `Status` empty, `Triaged = No` and `Task Created = No`. This is the triage queue.

**Phase 2 — Triage and tracking (Quality Team).** The Quality Manager opens the item in the app, classifies it (`Transformed into`), assigns a `Task Owner` and `Due Date`, and sets `Triaged = Yes`. From there the item is tracked through its status lifecycle until it is `Closed` or `Rejected`. The app reads and writes the list through Microsoft Graph; the reminder flow runs independently on a daily schedule.

```
Q-Star MS Form ──(flow: on form response)──▶ SharePoint list item  (Triaged=No)
                                                      │
                                  React app via Graph │  QM triage, owner work
                                                      ▼
                                              status lifecycle ──▶ Closed / Rejected
                                                      ▲
                  Daily reminder flow (recurrence) ───┘  reads list, sends Owner/QM/BU notices
```

---

## 2. SharePoint list — column reference

Create one list, suggested name **`Q-Star Issues`**, on your Quality site. Below, "Source" indicates whether the column already exists in your current list, or is **new** (added to support owner assignment and the NC effectiveness-test lifecycle).

A note on **internal names**: SharePoint derives the internal name from the display name at creation time (spaces become `_x0020_`). To keep Graph calls clean, create each column with the simple internal name shown below first, then rename the *display* name afterwards. The internal name is what you use in Graph `fields` objects and `$filter`.

### 2.1 Intake fields (already in your list)

| App field (`issue.*`) | Display name | Internal name | SharePoint type | Notes |
|---|---|---|---|---|
| `qsNumber` | Qs Number | `QsNumber` | Number | Sequential business key shown as `QS-{n}`. |
| `id` | ID | `ID` | Number | SharePoint's own item ID; do not create — use the built-in. |
| `shortSummary` | Short Summary | `ShortSummary` | Single line / multiline | One-line title. |
| `description` | Description | `Description` | Multiline (plain) | |
| `immediateAction` | Immediate Action taken | `ImmediateAction` | Multiline (plain) | |
| `severity` | Severity | `Severity` | Choice | Critical / High / Medium / Low. Drives the default due date. |
| `createdBy` | Created by | `ReportedBy` | Person *or* text | See §4.3 on person columns. Store the reporter's email if you want them notified. |
| `reportDate` | Report date | `ReportDate` | Date | |
| `departmentBU` | Department/Business Unit | `DepartmentBU` | Choice | 23 business units (see §3). |
| `region` | Region | `Region` | Choice | 7 regions. |
| `alreadyInContact` | Already in Contact | `AlreadyInContact` | Choice (Yes/No) | |
| `deviationType` | Deviation Type | `DeviationType` | Choice | 8 types. |
| `issueOrigin` | Origin | `Origin` | Choice | Customer Complaints or Claims / Internal Finding. |
| `attachments` | Attachment | (list attachments) | Attachment / library | See §4.4 — attachments are handled outside the `fields` object. |
| `additionalComments` | Additional Comments | `AdditionalComments` | Multiline (plain) | |

### 2.2 QM assessment fields (already in your list)

| App field | Display name | Internal name | Type | Notes |
|---|---|---|---|---|
| `followUp` | Follow up | `FollowUp` | Multiline (plain) | General QM notes. |
| `status` | Status | `Status` | Choice | **Add the new value** — see §2.4. |
| `transformedInto` | Transformed into | `TransformedInto` | Choice | OFI / NC Minor / NC Major / Only sent to Dept/BU for Action. |
| `taskCreated` | Task Created | `TaskCreated` | Choice (Yes/No) | Default `No`; set `Yes` when a tracked task is created at triage. |

### 2.3 New columns to add

These back the owner-assignment, escalation, §10.2 corrective-action and NC effectiveness-test features.

| App field | Display name | Internal name | Type | Notes |
|---|---|---|---|---|
| `triaged` | Triaged | `Triaged` | Choice (Yes/No) | Default `No`. Set `Yes` on the QM's first save. Distinguishes the triage queue from the register. |
| `taskOwner` | Task Owner | `TaskOwner` | Person | The named owner who receives reminders. Store/resolve email for notifications. |
| — | Permissioned Owner Email | `PermissionedOwnerEmail` | Single line text | Flow-maintained marker used to avoid reapplying item permissions when the owner has not changed. Not user-editable. |
| `ownerBU` | Escalation BU | `EscalationBU` | Choice | BU lead notified on 7-day overdue escalation. Defaults to `DepartmentBU`. |
| `dueDate` | Due Date | `DueDate` | Date | Auto-set from severity SLA at triage, QM-overridable. |
| `rootCause` | Root Cause | `RootCause` | Multiline | NC only (§10.2). |
| `correctiveAction` | Corrective Action | `CorrectiveAction` | Multiline | NC only. |
| `implementationDate` | Implementation Date | `ImplementationDate` | Date | NC only. **Start of the 2-month effectiveness test.** |
| `effectivenessCheck` | Effectiveness Check | `EffectivenessCheck` | Multiline | NC only. Evidence the fix held. |
| `verifiedBy` | Verified By | `VerifiedBy` | Person | NC only. Required before closing. |
| `verifiedDate` | Verified Date | `VerifiedDate` | Date | NC only. |
| `closedDate` | Closed Date | `ClosedDate` | Date | Stamped when status moves to Closed. |
| `closedAt` | Closed At | `ClosedAt` | DateTime | Full timestamp companion to `closedDate` (the app shows time-of-day when available). |
| `holdReason` | Hold Reason | `HoldReason` | Multiline (plain) | Set when status moves to On Hold. |
| `holdUntil` | Hold Until | `HoldUntil` | Date | Resume date; drives the "resumes in N days" reminder. |
| `ownerUpdate` | Owner Update | `OwnerUpdate` | Choice (Yes/No) | Default `No`. Set `Yes` when the Task Owner posts a comment or changes status, to flag the QM. |
| `ownerUpdateAt` | Owner Update At | `OwnerUpdateAt` | DateTime | Timestamp of the owner's last flagged update. |
| `ownerUpdateText` | Owner Update Text | `OwnerUpdateText` | Multiline (plain) | Short description of what the owner did (e.g. "Status changed from X to Y"). |

> These operational columns (`ClosedAt`, `HoldReason`, `HoldUntil`, `OwnerUpdate`, `OwnerUpdateAt`, `OwnerUpdateText`, `PermissionedOwnerEmail`) are included in both provisioning scripts.

### 2.4 Status choice — add the new value

The app introduces one new status used **only** by the NC effectiveness-test lifecycle. Add it to the `Status` choice column so the full set is:

```
Created
In Progress
Under Testing/Revision      ← NEW (NC only)
On Hold
Closed
Rejected
```

`Under Testing/Revision` is an active (open) status but is deliberately **excluded from overdue logic** — see §5.3.

### 2.5 Progress log — recommended as a child list

The app keeps an **append-only, timestamped** progress log (`{ ts, author, text }[]`). Two production options:

- **Recommended: a separate list `Q-Star Progress Log`** with columns `ParentItemId` (Number, lookup to the issue's `ID`), `Author` (Person/text), `EntryDate` (DateTime), `Text` (Multiline). Append-only is enforced by only ever creating items, never editing. The app fetches entries with `$filter=fields/ParentItemId eq {id}` ordered by `EntryDate`.
- **Simpler: a multiline column** `ProgressLog` with *Append Changes to Existing Text* enabled and list versioning on. SharePoint then stamps each addition with author and time automatically. This is easy but harder to render as structured entries.

The child list is closer to the app's model and is what the rest of this guide assumes.

---

## 3. Choice column allowed values

For reference when creating the choice columns (and when validating in the form/flow).

**Severity:** Critical, High, Medium, Low.

**Status:** Created, In Progress, Under Testing/Revision, On Hold, Closed, Rejected.

**Transformed into:** OFI, NC Minor, NC Major, Only sent to Dept/BU for Action. *(REC is reserved for forward-compatibility in the dashboard category chart but is not yet a taxonomy value — add it here and to `categoryOf` in the app if/when you adopt it.)*

**Deviation Type:** Communication, Compliance, Documentation, Equipment, Process, Quality, Safety, System.

**Origin:** Customer Complaints or Claims, Internal Finding.

**Region:** Americas (Miami), Asia Pacific, China (Shanghai), Eastern Europe (Vienna), Head Office (Neu-Isenburg), Western Europe (Amsterdam).

**Department/Business Unit (and Escalation BU):** BU Aftermarket, BU Airlines, BU Automotive, BU Diplo & High Security, BU High Tech & SemiCon, BU Life Science, Central Europe & Commercial Services, IT, Legal & Data Protection, Marketing, Network & Products, Quality, Risk Management, Strategy & Transformation, tmCT FRA, tmCT MUC, tmCT MEX/NLU, tmCT PVG. *(Extend to the full set used on your live form.)*

---

## 4. Data layer — implemented as SharePoint REST (PnPjs), not Graph

This replaces the app's `window.storage` calls. In the test build, the whole register lives under the storage key `qstar:issues:v2` and settings under `qstar:settings:v1`. In production, each issue is a SharePoint list item.

**Implementation note:** the web part runs *inside* the same SharePoint site as the lists (per the "single dedicated site" constraint), so the actual code (`frontend/src/webparts/qstarIssueManager/services/SharePointDataService.ts`) uses **SharePoint REST via PnPjs** (`@pnp/sp` with the `SPFx` behavior) instead of Microsoft Graph. This means:

- No Entra app registration, no admin consent, no `Sites.Selected` grant — the signed-in user's existing SharePoint permissions on the site are all that's needed.
- Same-tenant, same-site REST calls work out of the box; this is *more* least-privilege than the Graph approach below, not less.
- The production schema decision is to use native SharePoint Person columns (`ReportedBy`, `TaskOwner`, `VerifiedBy`, and progress-log `Author`). The current data service still treats some of these values as strings and unconditionally selects optional `*Email` columns; reconcile that implementation before tenant testing by expanding Person fields and writing lookup IDs.
- Settings (the IT-settings tab) are stored as one JSON blob in a single-item list, **`Q-Star Config`** (column `SettingsJson`), provisioned by both scripts.

The Graph-based data-access design below (§4.1–§4.5) is kept for reference in case the web part ever needs to reach a list on a *different* site than where it's hosted — that's the one scenario where Graph's cross-site reach earns its extra complexity.

### Connection diagnostics

`frontend/src/webparts/qstarIssueManager/services/ConnectionDiagnosticsService.ts` runs a live self-test from inside the web part: site access, current-user resolution, both lists exist with all expected columns, the `Under Testing/Revision` status choice is present, `Status`/`Triaged`/`DueDate` are indexed, and a full create → update → delete round-trip on a throwaway item. Surface this behind a button in the Admin/IT-settings tab so whoever deploys the web part in the real tenant can self-verify the connection without needing a developer to re-test it — this is the intended way to confirm "it works" against your actual SharePoint environment, since no external tool can authenticate as your tenant's users.

### 4.1 App registration (Entra ID)

Register a single-page application in Entra ID:

- **Platform:** SPA, redirect URI = your hosted app URL. Use **MSAL.js** with the authorization-code + PKCE flow.
- **Delegated scopes:** `User.Read`, `Sites.ReadWrite.All` *(or, for least privilege, `Sites.Selected` granted only on the Quality site — see §4.5)*, and `GroupMember.Read.All` to drive profile gating (§4.6).
- **Admin consent** for the tenant on the above scopes.

The reminder flow runs separately under its own connection/service identity (§5), not the user's token.

### 4.2 Resolving site and list IDs

Resolve once and cache:

```
GET /sites/{hostname}:/sites/Quality                      → siteId
GET /sites/{siteId}/lists?$filter=displayName eq 'Q-Star Issues'   → listId
```

### 4.3 Reading and writing items

List items carry their column data inside a `fields` object; always expand it.

```http
# Register (all triaged items) — maps to load of qstar:issues:v1
GET /sites/{siteId}/lists/{listId}/items?expand=fields&$top=2000

# Triage queue (untriaged)
GET /sites/{siteId}/lists/{listId}/items?expand=fields&$filter=fields/Triaged eq 'No'

# Create an item (Phase-1 intake done by the form flow, or app submit)
POST /sites/{siteId}/lists/{listId}/items
{ "fields": { "ShortSummary": "...", "Severity": "High", "Status": null, "Triaged": "No", "TaskCreated": "No", ... } }

# Update on QM save / owner update — maps to updateIssue()
PATCH /sites/{siteId}/lists/{listId}/items/{itemId}/fields
{ "Status": "Under Testing/Revision", "ImplementationDate": "2026-06-19", "TaskOwner": ... }
```

Filtering on non-indexed columns needs the header `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`, or index `Status`, `Triaged` and `DueDate` on the list (recommended for performance and for the reminder flow).

**Person columns** (`TaskOwner`, `VerifiedBy`, `ReportedBy`) are written by **LookupId**, not name. Resolve the user first (`GET /users/{email}` → `id`, then ensure they exist in the site user info list) and set `TaskOwnerLookupId`. Text plus companion `*Email` columns may remain in the isolated demo only; production uses native Person columns so permissions, reassignment, and notification identity are based on stable tenant users.

### 4.4 Attachments

Graph's support for classic SharePoint **list-item attachments** is limited. Two clean options:

- Store uploaded files in a **document library** and keep a hyperlink/`driveItem` reference column on the issue. This is the more Graph-native approach.
- Or keep classic list attachments and manage them via the SharePoint REST endpoint `/_api/web/lists/.../items({id})/AttachmentFiles` rather than Graph.

The test build treats attachments as a placeholder, so this is greenfield either way.

### 4.5 Least privilege with `Sites.Selected`

Rather than tenant-wide `Sites.ReadWrite.All`, register the app with `Sites.Selected` and have an admin grant write access to just the Quality site:

```http
POST /sites/{siteId}/permissions
{ "roles": ["write"], "grantedToIdentities": [{ "application": { "id": "{clientId}", "displayName": "Q-Star" } }] }
```

### 4.6 Profile gating from SharePoint and Entra groups

The app's Admin / Quality Manager / Task Owner switcher is a front-end simulation and must not ship as production authorization. The recommended same-site design is:

1. Create SharePoint groups named **Q-Star Admins**, **Q-Star Quality Managers**, **Q-Star Task Owners**, and **Q-Star Readers**.
2. Add users directly during the pilot. Where centralized governance is required, add the corresponding Entra security groups to those SharePoint groups.
3. Resolve one effective UI role from the signed-in user's SharePoint group membership/effective permissions, using `Admin > Quality Manager > Task Owner > Reader` priority.
4. Enforce the same access in SharePoint: Admin/QM edit all items, Readers read, and the assigned Task Owner receives item-level edit permission through Power Automate. UI tab visibility is not a security boundary.

Validate nested Entra-group behaviour in the real tenant. If SharePoint does not enumerate nested membership reliably enough for role selection, use SPFx's authenticated `MSGraphClientV3` only for the group-membership check and request the smallest delegated permission that IT approves. This still uses the existing Microsoft 365 sign-in; it does not require a custom login, client secret, or browser token store.

Keep the in-app role switcher only behind an explicit local-development flag. The production tab map follows the resolved role, fails closed while resolution is unavailable, and restricts the IT-settings tab to Admins.

### 4.7 Settings storage

The production settings object (form URL, flow ID, and any non-secret display/runtime options) can live in a small single-item **config list** or in app configuration. Remove the prototype's tenant/client-ID fields: same-site PnPjs uses the SPFx context, and an optional `MSGraphClientV3` role lookup also uses SPFx-managed authentication. The IT-settings tab is mostly read-only runtime configuration plus the editable Microsoft Form URL surfaced to reporters.

---

## 5. Reminder engine → Power Automate flow

The app's `buildReminders()` function simulates notifications. In production, reproduce it with a **scheduled cloud flow**. The rules below are exactly what the app implements, so behaviour stays identical.

### 5.1 Derived dates

- **Due date** is stored on the item (`DueDate`), set at triage from the severity SLA and QM-overridable. The default SLA, applied when the QM accepts the auto value: **Critical = 7, High = 14, Medium = 30, Low = 60** days from the report date.
- **Test end** (NC only) `= ImplementationDate + 2 months`. Compute in the flow with `addToTime(ImplementationDate, 2, 'Month')`.

### 5.2 Reminder rules

| # | Applies to | Trigger day | Recipient | Channel | Urgency |
|---|---|---|---|---|---|
| 1 | Open, not under test | `DueDate − 3` | Task Owner | Email/Teams | info (heads-up) |
| 2 | Open, not under test | `DueDate` | Task Owner | Email/Teams | warn (due today) |
| 3 | Open, not under test | every 3 days after `DueDate` (cap ~30d) | Task Owner | Email/Teams | danger (overdue nudge) |
| 4 | Overdue | `DueDate + 1` | Quality Team | Email/Teams | danger (QM alert) |
| 5 | Overdue ≥ 7d | `DueDate + 7` | Escalation BU lead | Email/Teams | danger (escalation) |
| 6 | NC under test | `TestEnd − 7` | Task Owner | Email/Teams | info (test ending soon) |
| 7 | NC under test | `TestEnd` | Task Owner **and** Quality Team | Email/Teams | warn (verify & close) |

### 5.3 Branching rule (important)

An item with `Status = Under Testing/Revision` is **open but not overdue**. The flow must:

- **Skip** rules 1–5 for it (no due-date chasing while the corrective action is being observed), and
- **Apply** rules 6–7 instead (effectiveness-check reminders around the test end date).

This matches `isOverdue()` in the app, which excludes the `Under Testing/Revision` status, and the `inNCTest()` branch in `buildReminders()`.

### 5.4 Flow structure

```
Recurrence (daily, e.g. 07:00 local)
└─ Get items: filter  Status ne 'Closed' and Status ne 'Rejected' and TaskCreated eq 'Yes'
   └─ Apply to each item
      ├─ Compose: daysToDue   = dateDifference(utcNow(), DueDate)
      ├─ Compose: testEnd     = addToTime(ImplementationDate, 2, 'Month')
      ├─ Condition: Status == 'Under Testing/Revision' ?
      │     ├─ YES → evaluate rules 6–7 against testEnd
      │     └─ NO  → evaluate rules 1–5 against DueDate
      ├─ Send Outlook email / Post Teams message to the matched recipient(s)
      └─ Upsert a Reminder Log entry (dedupe — see 5.5)
```

Use the **Office 365 Outlook** connector for email and/or **Microsoft Teams** (*Post message*) for chat notifications. Resolve recipients from the person columns (or the `*Email` text columns if you took the simpler route in §4.3).

### 5.5 De-duplication

A daily schedule naturally fires once per day, but you still need to avoid re-sending the same milestone. Two approaches:

- **Reminder Log list** keyed by `ItemId + RuleType + Date`: before sending, check whether a matching entry already exists today; create one after sending. Cleanest and auditable.
- **Last-notified field** on the item (`LastReminderType`, `LastReminderDate`): only send if the computed rule for today differs from what's recorded.

The overdue nudge (rule 3) is the every-3-days case — gate it on `mod(daysOverdue, 3) == 0` so it fires on day 3, 6, 9, … rather than daily.

### 5.6 Intake flow (Phase 1)

Separately, a second flow handles intake: trigger **When a new response is submitted** on the Q-Star form → *Get response details* → **Create item** in `Q-Star Issues` with `Status` empty, `Triaged = No`, `TaskCreated = No`, and `QsNumber` from your sequence (a counter item or `max(QsNumber)+1`). This is the production equivalent of the app's `addIntake()`.

---

## 6. Replication checklist

1. Run `backend/sharepoint/provisioning/provision-qstar-beta.ps1` (or `provision-qstar-beta-m365.sh`) for a beta with lists only. For production, run `provision-qstar.ps1` (or `provision-qstar-m365.sh`) to create the same native-Person list schema plus the four Q-Star role groups and permissions. Both profiles create `Q-Star Issues`, `Q-Star Progress Log`, and `Q-Star Config`, and index `Status`, `Triaged`, and `DueDate`.
2. ~~Register the Entra SPA app; grant `Sites.Selected`~~ — not needed with the implemented REST/PnPjs approach (§4). Only relevant if you later switch to Graph for a cross-site scenario.
3. ~~Swap the app's `window.storage` module for a Graph client~~ — done: `SharePointDataService.ts` implements load / create / update / progress-append against the lists via PnPjs. Component logic still needs to be ported from the prototype to actually call it (see `frontend/README.md`).
4. Deploy the web part to a page in the tenant (or run `gulp serve` against the tenant's SharePoint Workbench) and run the **connection diagnostics** (above) to self-verify site access, schema, and a write round-trip before doing anything else.
5. Wire the IT-settings tab values to `Q-Star Config`; surface the MS Form URL to reporters (already done in the app).
6. Create the four SharePoint role groups, implement role resolution and SharePoint item permissions (§4.6), and remove the production role switcher. Use a Graph membership check only if tenant testing requires it for nested Entra groups.
7. Build the **intake flow** (§5.6) and the **daily reminder flow** (§5.4) reproducing rules 1–7, including the under-test branch (§5.3) and dedupe (§5.5). These run as Power Automate cloud flows and can't be tested from inside the web part — see `backend/sharepoint/connection-test-plan.md` for the manual verification steps.
8. Verify against the seeded scenarios: an overdue task (rules 3–5), an NC mid-test (rule 6), and an NC whose test window has elapsed (rule 7, then QM verify-and-close unlocks).

---

*Field names, choice values, SLA days, the 2-month NC test period, and the reminder timings above are taken directly from the test build so the production system reproduces its behaviour one-for-one.*
