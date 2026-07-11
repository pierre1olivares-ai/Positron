# Q-Star Issue Manager

Quality-issue management tool for time:matters (Lufthansa Cargo group), built around ISO 9001:2015. Currently a working prototype, in progress of being productionized into an SPFx web part on SharePoint / Microsoft 365.

See [CLAUDE.md](CLAUDE.md) for full project context, domain notes, and next steps.

## Repo structure
- `src/` — prototype React app (`qstar-issue-manager.jsx`), starting point for the SPFx web part.
- `docs/` — clickable preview (`qstar-live.html`), SharePoint/Graph integration design, Power Automate flow docs, and the implementation checklist.
- `scripts/` — provisioning scripts for the SharePoint List columns (PowerShell + bash/M365 CLI variants).
