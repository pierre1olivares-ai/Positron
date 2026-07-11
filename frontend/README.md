# Frontend — Q-Star Issue Manager web part

**Current state:** `prototype/` holds the validated React prototype (`qstar-issue-manager.jsx`) and its clickable demo build (`qstar-live.html`). These are reference source only — not yet a buildable project.

**Target state:** an SPFx (SharePoint Framework) web part, scaffolded with the Yeoman generator, living at this `frontend/` root.

## Scaffolding the real project

Once Node.js LTS and the SPFx toolchain are installed (see repo root README), run from inside `frontend/`:

```bash
npm install -g yo gulp-cli @microsoft/generator-sharepoint
yo @microsoft/sharepoint
```

Answer the prompts as:
- Solution name: `qstar-issue-manager`
- Component type: WebPart
- Framework: React
- Component name: `QstarIssueManager`

This generates `package.json`, `tsconfig.json`, `gulpfile.js`, `config/`, and `src/webparts/qstarIssueManager/` alongside this README. From there, port the component logic out of `prototype/qstar-issue-manager.jsx` into `src/webparts/qstarIssueManager/components/QstarIssueManager.tsx`, and replace the `window.storage` calls with the Graph data layer described in [`../backend/sharepoint/qstar-sharepoint-graph-integration.md`](../backend/sharepoint/qstar-sharepoint-graph-integration.md).

Remove the Tailwind CDN `<script>` tag present in the prototype — production bundles all styling locally (see the hard constraints in [`../CLAUDE.md`](../CLAUDE.md)).
