# Q-Star Issue Manager

Quality-issue management tool for time:matters (Lufthansa Cargo group), built around ISO 9001:2015. Currently a working prototype, in progress of being productionized into an SPFx web part on SharePoint / Microsoft 365, sitting behind time:matters IT's standard Java backend template.

See [CLAUDE.md](CLAUDE.md) for full project context, domain notes, and next steps.

## Repo structure

```
frontend/         SPFx web part (React) — UI layer. See frontend/README.md.
backend/service/  Java/Spring Boot backend — thin authenticated gateway. See backend/service/README.md.
backend/          SharePoint List + Power Automate — the actual data/automation layer. See backend/README.md.
docs/             Cross-cutting project docs (rollout/implementation checklist).
```

- **`frontend/`** — the scaffolded SPFx web part. `frontend/prototype/` holds the original validated React prototype and its clickable demo, kept as reference until its UI is ported into the real web part. Its data layer can talk to SharePoint directly, or through the backend below — a web part property switches between them.
- **`backend/`** — SharePoint is the actual data store; there's no separate database of Q-Star business data. `backend/sharepoint/` documents the List schema and provisioning; `backend/power-automate/` handles intake/reminders; `backend/service/` is a Java gateway in front of SharePoint (validates the frontend's Azure AD tokens, applies business rules, proxies to SharePoint via Graph with its own app identity) — builds cleanly but isn't deployed anywhere yet.
- **`docs/`** — the non-developer rollout plan and IT ask list (`qstar-implementation-checklist.md`).

## Local development

**Frontend** — Node.js 18 LTS (via [nvm](https://github.com/nvm-sh/nvm), no admin rights needed) and the SPFx toolchain:

```bash
nvm install 18 && nvm use 18
npm install -g yo gulp-cli @microsoft/generator-sharepoint
cd frontend && npm install
npx gulp serve
```

See [frontend/README.md](frontend/README.md) for what's implemented (data layer, connection diagnostics) and what's left (porting the prototype UI). Before testing against your company's tenant, provision the SharePoint lists (`backend/sharepoint/provisioning/`) and follow [`backend/sharepoint/connection-test-plan.md`](backend/sharepoint/connection-test-plan.md).

**Backend** — JDK 21 (e.g. [Eclipse Temurin](https://adoptium.net/), no admin rights needed):

```bash
cd backend/service
./gradlew build
```

See [backend/service/README.md](backend/service/README.md) for what it needs to actually run (Postgres, Azure AD app registrations, SharePoint Graph access) and what's still needed from IT to deploy it (Azure DevOps, container registry access, Kubernetes).
