import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient } from "@microsoft/sp-http";
import { IResolvedRole, Role } from "../models/IRole";

export interface IBackendResponse {
  ok: boolean;
  status: number;
  statusText: string;
  headers: { get(name: string): string | null }; // eslint-disable-line @rushstack/no-new-null -- Fetch header contract.
  json(): Promise<unknown>;
}

export interface IBackendTransport {
  request(path: string, method?: string, body?: unknown, headers?: Record<string, string>): Promise<IBackendResponse>;
}

/** One authenticated transport is shared by data, role resolution and diagnostics. */
export class BackendApiClient implements IBackendTransport {
  public readonly baseUrl: string;

  constructor(private context: WebPartContext, private resourceId: string, baseUrl: string) {
    this.baseUrl = baseUrl.trim().replace(/\/+$/, "");
  }

  public async request(path: string, method: string = "GET", body?: unknown, headers: Record<string, string> = {}): Promise<IBackendResponse> {
    let url: URL;
    try { url = new URL(this.baseUrl); }
    catch { throw new Error("Configure a valid HTTPS backend base URL in the web part properties."); }
    if (url.protocol !== "https:" || url.username || url.password || url.search || url.hash || !this.resourceId.trim()) {
      throw new Error("Configure an HTTPS backend base URL and its Entra application ID or resource URI in the web part properties.");
    }
    const client = await this.context.aadHttpClientFactory.getClient(this.resourceId.trim());
    return client.fetch(this.baseUrl + path, AadHttpClient.configurations.v1, {
      method,
      headers: { Accept: "application/json", ...(body === undefined ? {} : { "Content-Type": "application/json" }), ...headers },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    });
  }
}

export async function backendError(response: IBackendResponse): Promise<Error> {
  let detail = "";
  try {
    const body = await response.json() as { title?: string; details?: string; message?: string };
    detail = body.message || body.title || body.details || "";
  } catch { /* Status still identifies the rejected request when its body is unreadable. */ }
  return new Error(`Q-Star backend request failed (${response.status}${response.statusText ? ` ${response.statusText}` : ""})${detail ? `: ${detail}` : "."}`);
}

export async function backendJson(client: IBackendTransport, path: string): Promise<unknown> {
  const response = await client.request(path);
  if (!response.ok) throw await backendError(response);
  return response.json();
}

/** /me is authoritative for both access and the server's configured data target. */
export function parseBackendSession(value: unknown, baseUrl: string): IResolvedRole {
  const session = value as {
    role?: string; source?: string;
    user?: { userId?: number; displayName?: string; email?: string };
    connection?: { siteUrl?: string; issuesListName?: string; progressListName?: string; betaAccessMode?: boolean };
  };
  const user = session && session.user;
  const connection = session && session.connection;
  if (!session || session.source !== "backend" || ["admin", "qm", "owner", "reader"].indexOf(session.role || "") < 0 ||
      !user || !positiveId(user.userId) || !user.displayName?.trim() || !user.email?.trim() ||
      !connection || !connection.siteUrl?.trim() || !connection.issuesListName?.trim() || !connection.progressListName?.trim() ||
      typeof connection.betaAccessMode !== "boolean") {
    throw new Error("The backend did not return a complete verified role, signed-in user and connection. Access is disabled.");
  }
  let site: URL;
  try { site = new URL(connection.siteUrl); }
  catch { throw new Error("The backend returned an invalid SharePoint site URL. Access is disabled."); }
  if (site.protocol !== "https:" || site.username || site.password || site.search || site.hash) {
    throw new Error("The backend returned an invalid SharePoint site URL. Access is disabled.");
  }
  return {
    role: session.role as Role, source: "backend",
    user: { userId: user.userId as number, displayName: user.displayName, email: user.email },
    connection: {
      siteUrl: connection.siteUrl, issuesListName: connection.issuesListName, progressListName: connection.progressListName,
      betaAccessMode: connection.betaAccessMode, dataSourceMode: "backend", backendBaseUrl: baseUrl,
    },
  };
}

export function positiveId(value: unknown): value is number {
  return typeof value === "number" && isFinite(value) && value > 0 && value <= 9007199254740991 && Math.floor(value) === value;
}
