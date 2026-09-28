# Q-Star Issue Manager

Quality-issue management tool for time:matters (Lufthansa Cargo group), built around ISO 9001:2015. Implemented as an SPFx web part on SharePoint / Microsoft 365. Tenant provisioning, automation deployment, and acceptance testing remain required before production use.

See the [implementation checklist](docs/qstar-implementation-checklist.md) for rollout responsibilities and [CLAUDE.md](CLAUDE.md) for project constraints.

## Working with issues

Admins and Quality Managers can submit reports in **Report**, then classify and assign them in **Triage queue**. Use the owner's Microsoft 365 name and email; changing only the email also changes the assigned identity. Task Owners work from **My tasks** and post notes with **Add update**. Saving detail fields and posting a progress note are separate actions.

For an NC, record the implementation date and start the **2-month effectiveness test**. Closing requires the issue to be under test, the two calendar months to have elapsed, and a verifier's name and Microsoft 365 identity. Both **Verify & close** and selecting **Closed** followed by **Save changes** enforce these rules. OFIs do not require the NC test or verifier fields. Putting an issue on hold requires a reason and a resume date of today or later.

Closed issues are read-only. Admins and Quality Managers can choose **Re-open issue** to return one to In Progress; an NC must complete a fresh test and verification cycle. The existing reference and journal remain. Dashboard year-to-date totals include today's reports using calendar dates.

## Saving and recovering drafts

- While a save, progress post, or reload is pending for an issue, its detail and progress controls are disabled, including after returning to the register and opening it again. Other issues remain editable. A close request cannot queue behind pending work: wait for it to finish, then initiate closing again.
- A rejected save or post keeps the open form's draft. A version conflict blocks another detail save until you reload. Copy any detail text you need before choosing **Reload latest and discard draft**; triage instead offers **Reload latest and return to queue**. Ordinary navigation is not an autosave.
- **Saved with a warning** means the write was accepted but a follow-up operation failed. Use the issue reload action instead of repeating the submission. An accepted closure stays read-only even if its refresh fails. If reopening is disabled because the saved version is unavailable, use **Reload this issue** in its read-only detail; it works independently of warnings for other issues and can be retried after a failure.
- If refreshed closure or reassignment makes an editor read-only, unsubmitted detail fields, hold fields, and progress text appear in **Unsaved draft recovery**. Unposted text is also retained when your own accepted close replaces the editor. These selectable copies remain separate from fresh edits, even after reopening, reassignment back, or later saves/posts. Repeated transitions can retain multiple copies. **Discard recovered draft** removes only that copy, without changing live edits or saved data.

Recovery copies are held only in the current app session. Copy anything you need before refreshing or closing the page or changing the web-part connection. Recovery never automatically restores or submits old content.

## Repo structure

```
frontend/   SPFx web part (React) — UI layer. See frontend/README.md.
backend/    SharePoint List + Power Automate — data/automation layer. See backend/README.md.
docs/       Cross-cutting project docs (rollout/implementation checklist).
```

- **`frontend/`** — the implemented SPFx web part. `frontend/prototype/` holds the original validated React prototype and its clickable demo, kept as a historical UI reference; the running implementation is under `frontend/src/`.
- **`backend/`** — no custom server; the SharePoint List is the data store, Power Automate handles intake/reminders. See `backend/sharepoint/` and `backend/power-automate/`.
- **`docs/`** — the non-developer rollout plan and IT ask list (`qstar-implementation-checklist.md`).

## Local development

Follow [frontend/README.md](frontend/README.md) for the supported Node version, installation, local development, checks, and packaging. For upgrade order, see [the repair handoff](docs/qstar-review-repairs.md). Before testing against your company's tenant, use the [provisioning guide](backend/sharepoint/provisioning/README.md) and [tenant verification plan](backend/sharepoint/connection-test-plan.md).
