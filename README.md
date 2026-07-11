# Q-Star Issue Manager

Quality-issue management tool for time:matters (Lufthansa Cargo group), built around ISO 9001:2015. Currently a working prototype, in progress of being productionized into an SPFx web part on SharePoint / Microsoft 365.

See [CLAUDE.md](CLAUDE.md) for full project context, domain notes, and next steps.

## Repo structure

```
frontend/   SPFx web part (React) — UI layer. See frontend/README.md.
backend/    SharePoint List + Graph + Power Automate — data/automation layer. See backend/README.md.
docs/       Cross-cutting project docs (rollout/implementation checklist).
```

- **`frontend/`** — the web part users interact with. `frontend/prototype/` holds the current validated React prototype and its clickable demo; this becomes the real SPFx solution once scaffolded.
- **`backend/`** — no custom server; the SharePoint List is the data store, Microsoft Graph is the API, Power Automate handles intake/reminders. See `backend/sharepoint/` and `backend/power-automate/`.
- **`docs/`** — the non-developer rollout plan and IT ask list (`qstar-implementation-checklist.md`).

## Local development

The frontend is not yet scaffolded as a runnable project (see [frontend/README.md](frontend/README.md) for the exact steps once Node.js and the SPFx toolchain are installed). Until then, `frontend/prototype/qstar-live.html` can be opened directly in a browser to demo the current UI.
