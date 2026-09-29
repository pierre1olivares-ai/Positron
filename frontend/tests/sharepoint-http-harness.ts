import { spfi } from "@pnp/sp";
import { DefaultHeaders, DefaultInit } from "@pnp/sp/behaviors/defaults";
import { BrowserFetchWithRetry, DefaultParse } from "@pnp/queryable";
import { SharePointDataService } from "../src/webparts/qstarIssueManager/services/SharePointDataService";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

type Row = Record<string, any>;
type Request = { method: string; path: string; url: string; body?: Row };
const site = "https://tenant.sharepoint.com/sites/quality";
const root = "/sites/quality/Lists/Progress";
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });

export function sharePointHttpHarness(initial: Row[] = [], relative = false) {
  const state = {
    issues: initial.map(row => ({ ...row })), progress: [] as Row[], requests: [] as Request[],
    reply: undefined as ((request: Request) => Response | Promise<Response>) | undefined,
    readReply: undefined as ((request: Request) => Response | undefined),
    failReads: false, failProgressRead: false, nextId: 21,
  };
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    const url = new URL(String(input));
    const path = decodeURIComponent(url.pathname);
    const method = init?.method || "GET";
    const body = init?.body ? JSON.parse(String(init.body)) : undefined;
    const request = { method, path, url: url.href, body };
    state.requests.push(request);
    const title = /getbytitle\('([^']+)'\)/i.exec(path)?.[1];
    const itemId = /\/items\((\d+)\)$/i.exec(path)?.[1];
    if (method === "GET") {
      const reply = state.readReply?.(request);
      if (reply) return reply;
      if (state.failReads) return json({ error: "Read unavailable" }, 403);
      if (path.endsWith("/_api/web")) return json({ Url: site });
      if (/getFolderByServerRelativePath/i.test(path)) return json({ Exists: true });
      if (title === "Q-Star Config") return json({ value: [{ Id: 1, ReferenceOffset: 3000 }] });
      if (title === "Q-Star Progress Log" && !/\/items/i.test(path)) return json({ RootFolder: { ServerRelativeUrl: root } });
      const rows = title === "Q-Star Issues" ? state.issues : state.progress;
      if (itemId) {
        if (title === "Q-Star Progress Log" && state.failProgressRead) return json({ error: "Read unavailable" }, 403);
        const row = rows.find(candidate => candidate.Id === Number(itemId));
        return json(row || { error: "Not found" }, row ? 200 : 404);
      }
      const folder = /FileDirRef eq '([^']+)'/.exec(url.searchParams.get("$filter") || "")?.[1];
      return json({ value: folder ? rows.filter(row => row.FileDirRef === folder) : rows });
    }
    if (method === "POST" && path.toLowerCase().endsWith("/ensureuser")) {
      return json({ Id: 7, Title: "Native user", EMail: body.logonName });
    }
    if (method === "POST" && title === "Q-Star Issues" && !itemId) {
      const saved = { ...body, Id: state.nextId++, "odata.etag": '"1"' };
      const response = state.reply ? await state.reply(request) : json(saved, 201);
      if (response.ok) state.issues.push(saved);
      return response;
    }
    if (method === "POST" && /addValidateUpdateItemUsingPath/i.test(path)) {
      const saved = {
        Id: 71 + state.progress.length, EntryText: body.formValues.find((field: Row) => field.FieldName === "EntryText").FieldValue,
        Author: { Id: 7, Title: "Native author", EMail: "native@example.com" }, Created: "2026-09-28T14:00:00Z",
        FileDirRef: new URL(body.listItemCreateInfo.FolderPath.DecodedUrl).pathname, FSObjType: 0,
      };
      const response = state.reply ? await state.reply(request) : json({ value: [{ FieldName: "Id", FieldValue: String(saved.Id), ItemId: saved.Id, HasException: false }] }, 201);
      if (response.ok) state.progress.push(saved);
      return response;
    }
    if (method === "POST" && title === "Q-Star Issues" && itemId) {
      Object.assign(state.issues.find(row => row.Id === Number(itemId))!, body);
      return new Response(null, { status: 204 });
    }
    throw new Error(`Unexpected request: ${method} ${path}`);
  };
  const client = spfi(site).using(DefaultInit(), DefaultHeaders(), BrowserFetchWithRetry({ interval: 1, retries: 2 }), DefaultParse());
  const service = relative ? new SharePointDataService({ pageContext: { web: { absoluteUrl: site } } } as WebPartContext) : new SharePointDataService({} as WebPartContext, undefined, undefined, undefined, client);
  return { service, state, restore: () => { globalThis.fetch = originalFetch; } };
}
