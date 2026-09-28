import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

import { IDataService } from "./IDataService";
import { IIssue, IProgressLogEntry } from "../models/IIssue";
import { ISettings, DEFAULT_SETTINGS } from "../models/ISettings";

/**
 * Production data layer once the Java backend (backend/service/) is deployed: calls its REST API
 * over an Azure-AD-secured connection (SPFx's AadHttpClient) instead of talking to SharePoint
 * directly. The backend is a thin gateway — SharePoint is still the actual store, reached
 * server-side via Microsoft Graph with the backend's own app identity — so from this class's
 * point of view it's just a normal JSON REST API. Field names match
 * backend/service/api-contract/contract.yaml's Issue/Settings/ProgressLogEntry schemas 1:1 with
 * IIssue.ts, so responses need no remapping.
 */
export class BackendApiDataService implements IDataService {
  constructor(
    private context: WebPartContext,
    private backendResourceId: string,
    private backendBaseUrl: string
  ) {}

  public async loadIssues(): Promise<IIssue[]> {
    const response = await this.get("/issues");
    return (await response.json()) as IIssue[];
  }

  public async createIssue(issue: Partial<IIssue>): Promise<IIssue> {
    const response = await this.send("/issues", "POST", issue);
    return (await response.json()) as IIssue;
  }

  public async updateIssue(id: number, patch: Partial<IIssue>): Promise<void> {
    await this.send(`/issues/${id}`, "PATCH", patch);
  }

  public async addProgressLogEntry(id: number, entry: IProgressLogEntry): Promise<void> {
    await this.send(`/issues/${id}/progress`, "POST", entry);
  }

  public async loadSettings(): Promise<ISettings> {
    try {
      const response = await this.get("/settings");
      return { ...DEFAULT_SETTINGS, ...((await response.json()) as Partial<ISettings>) };
    } catch {
      return DEFAULT_SETTINGS;
    }
  }

  public async saveSettings(settings: ISettings): Promise<void> {
    await this.send("/settings", "PUT", settings);
  }

  private async getClient(): Promise<AadHttpClient> {
    return this.context.aadHttpClientFactory.getClient(this.backendResourceId);
  }

  private async get(path: string): Promise<HttpClientResponse> {
    const client = await this.getClient();
    const response = await client.get(this.backendBaseUrl + path, AadHttpClient.configurations.v1);
    return this.checkOk(response);
  }

  private async send(path: string, method: string, body: unknown): Promise<HttpClientResponse> {
    const client = await this.getClient();
    const response = await client.fetch(this.backendBaseUrl + path, AadHttpClient.configurations.v1, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    return this.checkOk(response);
  }

  private async checkOk(response: HttpClientResponse): Promise<HttpClientResponse> {
    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(`Q-Star backend request failed: ${response.status} ${response.statusText} ${text}`);
    }
    return response;
  }
}
