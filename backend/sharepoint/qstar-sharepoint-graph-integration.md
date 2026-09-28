# Q-Star — SharePoint integration contract

The deployed SPFx web part uses SharePoint REST through PnPjs with the signed-in user's same-site permissions. No separate Graph data layer or custom API server is involved. Microsoft Forms intake, scheduled notifications, and assignment ACL reconciliation run as three tenant Power Automate flows under a service identity in Q-Star Admins.

This document describes the implemented schema. [Provisioning and migrations](provisioning/README.md) and the [workflow guide](../power-automate/qstar-power-automate-flows.md) are the operational instructions. The standalone root JSX prototype is not the production data service.

## 1. Data flow

Forms or an authorized app user creates an untriaged issue. The Quality Manager classifies it, assigns the owner/due date, and creates the tracked task. SharePoint stores lifecycle state and the separate secured journal. The assignment flow reconciles issue/folder access, and the daily dispatcher sends date milestones and accepted journal/status events.

## 2. SharePoint list — column reference

Provision **Q-Star Issues**, **Q-Star Progress Log**, and **Q-Star Config** on the same site as the web part. Use the [provisioning entry points](provisioning/README.md), including the explicit migration modes for an existing register.

A note on **internal names**: SharePoint derives the internal name from the display name at creation time (spaces become `_x0020_`). To keep REST calls consistent, create each column with the simple internal name shown below first, then rename the *display* name afterwards. The internal name is what you use in REST payloads and `$filter`.

### 2.1 Intake fields (already in your list)

| App field (`issue.*`) | Display name | Internal name | SharePoint type | Notes |
|---|---|---|---|---|
| `qsNumber` | Qs Number | `QsNumber` | Number | Stable business reference shown as `QS-{n}`. New references use Config.ReferenceOffset + ID; preserve existing nonempty values. |
| `id` | ID | `ID` | Number | SharePoint's own item ID; do not create — use the built-in. |
| `shortSummary` | Short Summary | `ShortSummary` | Single line / multiline | One-line title. |
| `description` | Description | `Description` | Multiline (plain) | |
| `immediateAction` | Immediate Action taken | `ImmediateAction` | Multiline (plain) | |
| `severity` | Severity | `Severity` | Choice | Critical / High / Medium / Low. Drives the default due date. |
| `createdBy` | Created by | `ReportedBy` | Native Person | Expanded for reads; written by SharePoint lookup ID. |
| `reportDate` | Report date | `ReportDate` | Date | |
| `departmentBU` | Department/Business Unit | `DepartmentBU` | Choice | 23 business units (see §3). |
| `region` | Region | `Region` | Choice | 6 canonical regions; reviewed migration preserves custom historical values. |
| `alreadyInContact` | Already in Contact | `AlreadyInContact` | Choice (Yes/No) | |
| `deviationType` | Deviation Type | `DeviationType` | Choice | 8 types. |
| `issueOrigin` | Origin | `Origin` | Choice | Customer Complaints or Claims / Internal Finding. |
| `attachments` | Attachment | (list attachments) | Attachment / library | Upload and attachment retrieval are not implemented by the current data service. |
| `additionalComments` | Additional Comments | `AdditionalComments` | Multiline (plain) | |

### 2.2 QM assessment fields (already in your list)

| App field | Display name | Internal name | Type | Notes |
|---|---|---|---|---|
| `followUp` | Follow up | `FollowUp` | Multiline (plain) | General QM notes. |
| `status` | Status | `Status` | Choice | Full set in §2.4. |
| `transformedInto` | Transformed into | `TransformedInto` | Choice | OFI / NC Minor / NC Major / Only sent to Dept/BU for Action. |
| `taskCreated` | Task Created | `TaskCreated` | Choice (Yes/No) | Default `No`; set `Yes` when a tracked task is created at triage. |

### 2.3 New columns to add

These back the owner-assignment, escalation, §10.2 corrective-action and NC effectiveness-test features.

| App field | Display name | Internal name | Type | Notes |
|---|---|---|---|---|
| `triaged` | Triaged | `Triaged` | Choice (Yes/No) | Default `No`. Set `Yes` on the QM's first save. Distinguishes the triage queue from the register. |
| `taskOwner` | Task Owner | `TaskOwner` | Person | The named owner who receives reminders. Store/resolve email for notifications. |
| — | Permissioned Owner Email | `PermissionedOwnerEmail` | Single line text | Legacy compatibility field; never use this writable value to skip ACL reconciliation. |
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
| `ownerUpdate` | Owner Update | `OwnerUpdate` | Choice (Yes/No) | Default `No`. UI acknowledgement flag for owner status updates. Progress comments are separate durable events. |
| `ownerUpdateAt` | Owner Update At | `OwnerUpdateAt` | DateTime | UTC timestamp of the latest owner status update; notification identity must not depend on the badge Boolean. |
| `ownerUpdateText` | Owner Update Text | `OwnerUpdateText` | Multiline (plain) | Short description of what the owner did (e.g. "Status changed from X to Y"). |

| `reminderCycle` | Reminder Cycle | `ReminderCycle` | Single line text | Starts as `initial`; a new token on reopening permits a fresh reminder cycle. |

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

`Under Testing/Revision` is an active (open) status but is deliberately **excluded from overdue logic** — see the flow guide.

### 2.5 Secured progress journal

**Q-Star Progress Log** is a generic list with folders and content types enabled. Each issue has one folder named `issue-<ID>` under the list's actual RootFolder.ServerRelativeUrl. The service derives the issue from server-managed FileDirRef; ParentItemId is optional legacy metadata and is not trusted for grouping.

Entries use EntryText (required plain multiline), built-in Author and Created. EntryDate remains a compatibility column. Append uses SharePoint AddValidateUpdateItemUsingPath with the absolute issue-folder URL; the caller cannot override Author/Created. Existing root-level rows are migrated in place after preview, preserving item ID and author/time. Invalid or orphan parent mappings stop migration for review.

In production, the list root grants Admin Full, QM **Q-Star Append Progress**, and Task Owners/Readers Read. Each issue folder adds only its current owner to Append. The custom role contains Read and Add Items without Edit/Delete/Manage Permissions. Folder inheritance protects child entries. Admins retain maintenance access; this is not an immutable compliance archive. Beta retains existing site permissions and therefore does not enforce append-only history.

---

## 3. Choice column allowed values

For reference when creating the choice columns (and when validating in the form/flow).

**Severity:** Critical, High, Medium, Low.

**Status:** Created, In Progress, Under Testing/Revision, On Hold, Closed, Rejected.

**Transformed into:** OFI, NC Minor, NC Major, Only sent to Dept/BU for Action. *(REC is reserved for forward-compatibility in the dashboard category chart but is not yet a taxonomy value — add it here and to `categoryOf` in the app if/when you adopt it.)*

**Deviation Type:** Communication, Compliance, Documentation, Equipment, Process, Quality, Safety, System.

**Origin:** Customer Complaints or Claims, Internal Finding.

**Region:** Americas (Miami), Asia Pacific, China (Shanghai), Eastern Europe (Vienna), Head Office (Neu-Isenburg), Western Europe (Amsterdam).

**Department/Business Unit (and Escalation BU):** BU Aftermarket, BU Airlines, BU Automotive, BU Diplo & High Security, BU High Tech & SemiCon, BU Life Science, Central Europe & Commercial Services, Claims & Complaints, Customer Solution & Business Development, Digital Transformation & Data Management, Finance & Controlling, Human Resources, IT, Legal & Data Protection, Marketing, Network & Products, Quality, Risk Management, Strategy & Transformation, tmCT FRA, tmCT MUC, tmCT MEX/NLU, tmCT PVG. *(Extend to the full set used on your live form.)*

---

## 4. Service behavior

- Native Person reads expand identity fields. Writes resolve the tenant user and use lookup IDs; text-plus-email schemas are rejected by provisioning.
- ReferenceOffset is a numeric Config field initialized once to at least the greatest legacy QsNumber (minimum 1000). Create the issue first, then use ReferenceOffset + its SharePoint ID. Preserve every nonempty old reference; do not recalculate the offset. A failed number-materialization write must resume the existing item, not repeat intake.
- Config contains one settings item. SettingsJson holds non-secret runtime settings; ReferenceOffset is separate and immutable after first use. Production users can read Config; only Admins can write it.
- Progress appends target the issue folder and use the server-created item and identity. A successful comment needs no separate issue update. QMs/Admins may create a missing folder only when their SharePoint permissions permit; an owner with root Read waits for Flow C.
- Date-only business values preserve their `yyyy-MM-dd` component. Event timestamps are UTC. Month addition clamps month ends; for example 31 December + two months is the last day of February.
- Production role resolution controls presentation; SharePoint ACLs enforce access. SharePoint Edit is item-level, not column-level authorization. Reassignment is asynchronous; former access persists until Flow C successfully removes it.
- Same-site list attachments are not implemented by the journal or issue service. Do not promise attachment upload based on the standalone prototype.

## 5. Notifications

Use [Flow B](../power-automate/qstar-power-automate-flows.md#b--reminders-and-accepted-update-notifications) as the single workflow specification. It branches NC testing, On Hold, and normal active tasks before evaluating dates. Hold and testing suppress normal due/overdue notices. Missing branch dates create an operational alert instead of evaluating null date expressions.

Dedupe includes event/cycle, milestone date, and the actual recipient. Accepted journal entries use their item ID. Owner status snapshots use OwnerUpdateAt and do not depend on the dismissible OwnerUpdate flag. A service-owned Pending/Sent log makes ambiguous send outcomes visible for reconciliation. The repository contains instructions, not exported or deployed flows.

## 6. Deployment verification

Run the [connection and workflow test plan](connection-test-plan.md) with Admin, QM, two distinct owners, and Reader accounts. Local stubs validate provisioning behavior but cannot validate tenant permissions, installed PnP versions, Power Automate expressions, or delivery. Complete those checks before enabling production writes.
