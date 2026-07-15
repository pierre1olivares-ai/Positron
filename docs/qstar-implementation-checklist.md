# Q‑Star Issue Manager — Implementation Guide & IT Checklist

**From prototype to a fully working tool in your Microsoft 365 / SharePoint environment.**

This guide is written for a **non‑developer leading the project**. Your job is to *coordinate and decide*; IT will *execute* the technical parts. Tick the boxes as you go, and use the "Who to ask in IT" section to send the right request to the right team.

---

## 0. The big picture — what we're building and where it lives

Think of Q‑Star as four pieces that work together:

| Piece | What it is | Where it lives |
|---|---|---|
| **The app** | The screen people use (the React tool you've been reviewing) | A page on a dedicated SharePoint site |
| **The data** | Every issue/task and its history | A **SharePoint List** on that same site — this is your "repository" of records |
| **Identity** | Who is Admin / Quality Manager / Task Owner / Reader | Microsoft 365 sign‑in (Entra ID) |
| **Automations** | The email reminders and the intake form | Power Automate + a Microsoft Form |

```
   Microsoft Form  ──►  Power Automate  ──►  SharePoint List (the data)  ◄──►  Q-Star app (the screen)
   (report a new issue)   (intake + daily               │                         ▲
                           reminder emails)              │                         │
                                                         └──── Microsoft 365 sign-in decides the role
                                                              (Admin / QM / Task Owner / Reader)
```

**The containment rule you asked for** — *"the tool must not have access outside the repository"* — is achieved by hosting the app **inside one dedicated SharePoint site** and giving it permission to touch **only that one site**, nothing else in the company. The exact controls are in **Section 4**, which is the part to show IT Security.

> **A quick note on the word "repository."** In software it usually means a **Git repository** — the place where the app's *source code* is stored and version‑controlled. In your sentence it most likely also means **the SharePoint site** where the app and its data live. Both matter, and this guide covers both: the **code** goes into a Git repository owned by IT Development (Section 2, Phase B), and the **running tool + data** are contained to one SharePoint site (Phase C–E). "No access outside the repository" = the running tool can only reach its own SharePoint site.

---

## 1. Five decisions to make first (with my recommendation)

Make these calls *before* IT starts, because everything else depends on them.

- [ ] **1.1 How the app is hosted.**
  **Recommended: build it as an "SPFx web part" that runs inside SharePoint.** SPFx (SharePoint Framework) is Microsoft's official way to run custom React code inside SharePoint Online. It runs under the user's existing sign‑in, needs no separate server, and is naturally contained to the site. *Alternative:* host it as a separate web app and call Microsoft Graph — more moving parts and a bigger security review, so only choose this if IT prefers it.

- [ ] **1.2 Where the data lives.**
  **Recommended: create one brand‑new, dedicated SharePoint site** (e.g. "Q‑Star — Business Excellence") that contains the List and the app page. A dedicated site keeps permissions clean and makes the "no outside access" promise easy to prove.

- [ ] **1.3 How roles are assigned.**
  **Recommended: four SharePoint site groups** (Q‑Star Admins, Q‑Star Quality Managers, Q‑Star Task Owners, Q‑Star Readers), backed by corresponding Entra ID security groups where central governance is required. The app uses the existing Microsoft 365 sign-in and resolves a role from site membership/effective permissions—not from a hand-kept email list. The assigned Task Owner receives edit permission only on their assigned issue.

- [ ] **1.4 Data classification.**
  Quality issues can include customer‑complaint details. Ask IT Security / your Data Protection contact to **classify the data** (almost certainly "Internal", possibly with personal data inside complaints) so the right retention and access rules apply.

- [ ] **1.5 Who owns it long‑term.**
  Name a **product owner** (likely you) and ask IT Development to name a **technical owner** who maintains the code after go‑live.

---

## 2. The phased plan — your master checklist

Each phase lists *what happens* and *who does it*. You drive; IT builds.

### Phase A — Governance & sign‑off *(you)*
- [ ] Confirm Q‑Star is an officially sponsored tool (you report to the CEO — get that sponsorship in writing).
- [ ] Share the live preview (`qstar-live.html`) and the QM feedback you've collected with IT as the requirements baseline.
- [ ] Agree the five decisions in Section 1 with IT.
- [ ] Open a ticket / project request with each IT team (Section 3).

### Phase B — Source control: the Git repository *(IT Development)*
- [ ] IT Development creates a **private Git repository** (Azure DevOps or GitHub Enterprise — whatever your company uses) for the app's code.
- [ ] They commit the current code (`qstar-issue-manager.jsx`) as the starting point.
- [ ] They set up branch protection and code review so changes are controlled.
- [ ] *(Optional but recommended)* a build pipeline so a reviewed change automatically produces the deployable package.

### Phase C — SharePoint site & List: the data home *(IT Infra / SharePoint admin)*
- [ ] Create the dedicated SharePoint site (decision 1.2).
- [ ] Create the **Issues List** with the correct columns. For beta use `provision-qstar-beta.ps1` or `provision-qstar-beta-m365.sh`; these create lists only. The production scripts additionally create the four Q-Star groups.
- [ ] Confirm the column names/types match the app's field map in `qstar-sharepoint-graph-integration.md`.
- [ ] Set site membership: only people who should use Q‑Star are members.

### Phase D — Identity & permissions: the containment *(IT Security / Entra admin)*
- [ ] **Beta:** enable the web part's Beta access mode and use the Communication site's existing Owners, Members, and Visitors permissions; provisioning skips Q-Star group creation by default.
- [ ] **Production:** create the four SharePoint role groups (decision 1.3), add users, and nest the corresponding Entra security groups where required by IT governance.
- [ ] Configure list permissions for Admin/QM edit and Reader read, plus item-level edit for the assigned Task Owner. Automate grant/revoke on assignment changes.
- [ ] Validate nested Entra-group role resolution in the tenant. Only if needed, approve the smallest delegated Microsoft Graph group-membership permission for the SPFx solution.
- [ ] If any Microsoft Graph **site data** access is later needed, register the app and grant **`Sites.Selected`** — restricted to **only the Q‑Star site** (see Section 4). It is not needed for the current same-site PnPjs data layer.
- [ ] Confirm **no tenant‑wide permissions** are granted.
- [ ] Apply your standard Conditional Access (MFA etc.) to the site.

### Phase E — Package & deploy the app *(IT Development + SharePoint admin)*
- [x] The validated UI is implemented as an **SPFx web part** and wired to SharePoint through PnPjs with native Person fields, paging, diagnostics, and SharePoint-group role resolution.
- [x] Tailwind, icons, and charts are bundled into the solution; the production web part does not load UI assets from a public CDN.
- [x] A clean production package is available at `frontend/sharepoint/solution/qstar-issue-manager.sppkg`.
- [ ] SharePoint admin uploads the package (`.sppkg`) to the **App Catalog** and approves it.
- [ ] Add the Q‑Star web part to a page on the dedicated site.

### Phase F — Automations: form + reminders *(IT Infra / Power Platform)*
- [ ] Build the **intake Microsoft Form** (the "Report an issue" entry point).
- [ ] Build the **three Power Automate flows** — intake, assignment permissions, and reminders/notifications. **You have the step‑by‑step build guide**: `qstar-power-automate-flows.md`. This is where assignment-level access and the owner-comment/status-change emails to Quality Managers are enforced.
- [ ] Confirm reminder emails come from an approved mailbox or service account.

### Phase G — Test it properly (UAT) *(you + a few pilot users)*
- [ ] Test each role end‑to‑end: Admin, Quality Manager, Task Owner, Reader.
- [ ] Test the full lifecycle: report → triage → assign → progress → on‑hold → effectiveness test (NC) → close → re‑open.
- [ ] Confirm reminders and the QM notifications actually arrive by email.
- [ ] Confirm a Reader sees **only the Dashboard** and cannot reach the register.
- [ ] Confirm the app cannot reach any other SharePoint site (IT Security validates).

### Phase H — Go live, train, support *(you + IT)*
- [ ] Short training for QMs and Task Owners (you can reuse your QM email + the live preview).
- [ ] Write a one‑page "how to use Q‑Star" for staff.
- [ ] Agree a support model: who fixes bugs, how change requests are raised.
- [ ] Switch the entities over, and retire any spreadsheet/old process it replaces.

---

## 3. Who to ask in IT — and exactly what to ask

| IT team | What they own | What to ask them for |
|---|---|---|
| **IT Infrastructure / Microsoft 365 & SharePoint admin** | SharePoint sites, the App Catalog, Microsoft Forms, Power Automate environment, service mailboxes | "Please create a dedicated SharePoint site for Q‑Star, create the Issues List using this provisioning script, deploy the supplied App Catalog package, and set up the intake Form and the three Power Automate flows from this guide." |
| **IT Security / Identity (Entra ID / Azure AD)** | Sign‑in, security groups, app registrations & permissions, data classification, Conditional Access | "Please approve four Q‑Star SharePoint role groups backed by Entra security groups, validate nested membership, and confirm that permissions are enforced only on the Q‑Star site and assigned items. No tenant-wide permissions or custom login should be introduced." |
| **IT Development** | The Git repository, packaging the app (SPFx), the SharePoint/Graph data layer, code review, long‑term maintenance | "Please put this React code into a managed Git repo, convert it to an SPFx web part that stores data in the SharePoint List (per this integration spec), bundle all assets locally with no external internet calls, and own the build/deploy." |
| **(If you have one) Data Protection Officer / Compliance** | Personal‑data handling, retention | "Customer complaints may contain personal data — please confirm classification, retention, and any DPIA need." |

> Tip: send each team the **specific deliverable** they need (next section). A concrete script or spec turns a vague request into a quick task.

---

## 4. The "no access outside the repository" guarantee *(show this to IT Security)*

This is the part that proves the tool is sandboxed. Ask IT Security to confirm each control:

- [ ] **One dedicated site collection.** The app and its List live in a single SharePoint site; nothing else is in it.
- [ ] **Site‑scoped permissions only.** If the app uses Microsoft Graph, it is granted **`Sites.Selected`** with access to **only the Q‑Star site by name** — the modern Microsoft control that prevents an app from reaching any other site. No `Sites.Read.All` / `Sites.FullControl.All` / tenant‑wide scopes.
- [ ] **Runs as the signed‑in user.** With SPFx, the app acts under each user's own Microsoft 365 permissions — so it can never see more than the person already can.
- [ ] **No secrets in the browser.** No passwords or API keys live in the front‑end code.
- [ ] **No public‑internet calls.** All code libraries are **bundled inside the package**; the production app must not load anything from external websites/CDNs. *(The preview build does use one external style library — that must be removed for production.)*
- [ ] **Data stays in the tenant.** All data lives in your SharePoint/Microsoft 365 tenant; nothing is sent to any outside service.
- [ ] **Least privilege on the List.** List/site access is limited to the four role groups; assigned owners receive edit only on their own items, with grants revoked on reassignment.

If all boxes are ticked, the tool is provably contained to its own repository/site.

---

## 5. What you already have to hand to IT (assets that save them days)

You're not starting from zero — give IT these existing files:

| File | Give it to | Why it helps |
|---|---|---|
| `qstar-issue-manager.jsx` | IT Development | The full working app — their starting code |
| `qstar-live.html` | Everyone | A clickable preview to demo and to confirm requirements |
| `qstar-sharepoint-graph-integration.md` | IT Development | The exact List columns and the read/write data layer |
| `provision-qstar-beta.ps1` *or* `provision-qstar-beta-m365.sh` | SharePoint admin | Beta: creates lists and columns only |
| `provision-qstar.ps1` *or* `provision-qstar-m365.sh` | SharePoint admin | Production: creates lists, columns, role groups, and site permissions |
| `qstar-power-automate-flows.md` | Power Platform / Infra | Step‑by‑step build of the intake + reminder flows |
| This checklist | You + all IT teams | The overall plan and ownership |

---

## 6. Order of events (what must happen before what)

1. **Decisions** (Section 1) — you, with IT.
2. **Git repo** set up — IT Development. *(can run in parallel with 3)*
3. **SharePoint site + List** created — IT Infra.
4. **SharePoint role groups + Entra membership + item permissions** — IT Security. *(needs the site from step 3)*
5. **App converted to SPFx + deployed** — IT Development + admin. *(needs steps 2–4)*
6. **Form + Power Automate flows** — IT Infra. *(needs the List from step 3)*
7. **UAT** — you + pilot users. *(needs steps 5–6)*
8. **Go live + training** — you + IT.

Steps 2–4 and 6 can largely overlap; 5 and 7 are the gates.

---

## 7. A realistic word on effort and your role

- **Your role:** product owner — requirements, decisions, testing (UAT), training, and governance. You do **not** need to write code.
- **IT Development effort:** typically **one developer for a number of days** to convert the prototype to SPFx, wire it to the List, and bundle assets — plus the flows. The prototype and the integration spec remove most of the guesswork, but converting a single‑file app to a production SPFx solution is genuine engineering work, not a copy‑paste.
- **Biggest dependencies:** getting the SharePoint site + permissions (IT Security sign‑off) and a developer assigned. Start those conversations first.
- **Keep the loop short:** the in‑app **Feedback** button and your QM feedback round give you a tested requirements set — share it so IT builds the right thing once.

---

## 8. Plain‑language glossary

- **SharePoint List** — a structured table inside SharePoint; here it stores every issue/task. Your "data repository."
- **SPFx (SharePoint Framework)** — Microsoft's official toolkit for running custom React code *inside* SharePoint Online.
- **Entra ID (formerly Azure AD)** — Microsoft 365's identity system; it knows who each person is when they sign in.
- **Security group** — a managed list of people in Entra ID; you grant access to the group, not person‑by‑person.
- **App Catalog** — the SharePoint store your admins use to approve and deploy custom apps (`.sppkg` packages).
- **`.sppkg`** — the packaged app file IT uploads to the App Catalog.
- **Microsoft Graph** — the Microsoft 365 API for reading/writing data across the tenant.
- **`Sites.Selected`** — a Graph permission that limits an app to **only specific, named SharePoint sites** — the control that delivers "no access outside the repository."
- **Power Automate** — Microsoft 365's automation tool; here it sends the reminder and notification emails.
- **Microsoft Forms** — the simple form used to report a new issue.
- **CDN** — a public‑internet location that serves code libraries; for containment, the production app must **not** use one.
- **Git repository** — the version‑controlled store for the app's source code.
- **UAT (User Acceptance Testing)** — structured testing by real users before go‑live.
- **DPIA** — a Data Protection Impact Assessment, sometimes needed when handling personal data.

---

*Prepared for the Q‑Star Issue Manager rollout · time:matters Business Excellence.*
