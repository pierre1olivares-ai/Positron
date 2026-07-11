# Q-Star Issue Manager

Quality-issue management tool for time:matters (Lufthansa Cargo group), built around ISO 9001:2015. Currently a working prototype, in progress of being productionized into an SPFx web part on SharePoint / Microsoft 365.

See [CLAUDE.md](CLAUDE.md) for full project context, domain notes, and next steps.

## Repo structure

```
frontend/   SPFx web part (React) — UI layer. See frontend/README.md.
backend/    SharePoint List + Power Automate — data/automation layer. See backend/README.md.
docs/       Cross-cutting project docs (rollout/implementation checklist).
```

- **`frontend/`** — the scaffolded SPFx web part. `frontend/prototype/` holds the original validated React prototype and its clickable demo, kept as reference until its UI is ported into the real web part.
- **`backend/`** — no custom server; the SharePoint List is the data store, Power Automate handles intake/reminders. See `backend/sharepoint/` and `backend/power-automate/`.
- **`docs/`** — the non-developer rollout plan and IT ask list (`qstar-implementation-checklist.md`).

## Local development

Requires Node.js 18 LTS (installed here via [nvm](https://github.com/nvm-sh/nvm) — no admin/sudo rights needed, unlike Homebrew) and the SPFx toolchain:

```bash
nvm install 18 && nvm use 18
npm install -g yo gulp-cli @microsoft/generator-sharepoint
cd frontend && npm install
npx gulp serve
```

See [frontend/README.md](frontend/README.md) for what's implemented so far (data layer, connection diagnostics) and what's left (porting the prototype UI). Before testing against your company's tenant, provision the SharePoint lists (`backend/sharepoint/provisioning/`) and follow [`backend/sharepoint/connection-test-plan.md`](backend/sharepoint/connection-test-plan.md).
