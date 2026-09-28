import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient } from "@microsoft/sp-http";

import { ICheckResult } from "./ConnectionDiagnosticsService";

/**
 * Frontend-to-backend connection self-test: calls the Java backend's own GET /diagnostics
 * (DiagnosticsService.java / DiagnosticsApiController.java), which checks the backend's Graph
 * app-identity against SharePoint. A successful call here proves two things at once: the SPFx
 * web part can obtain an Azure AD token for the backend and reach it (this class's own concern),
 * and — via the returned checks — that the backend can in turn reach SharePoint. Complements
 * ConnectionDiagnosticsService.ts, which checks the browser's own direct SharePoint access
 * instead (relevant only if SharePointDataService.ts / PnPjs is what's actually wired in).
 */
export class BackendDiagnosticsService {
  constructor(
    private context: WebPartContext,
    private backendResourceId: string,
    private backendBaseUrl: string
  ) {}

  public async run(): Promise<ICheckResult[]> {
    try {
      const client = await this.context.aadHttpClientFactory.getClient(this.backendResourceId);
      const response = await client.get(this.backendBaseUrl + "/diagnostics", AadHttpClient.configurations.v1);
      if (!response.ok) {
        return [
          {
            name: "Backend connection",
            status: "fail",
            message: `Reached the backend but it returned ${response.status} ${response.statusText}.`,
          },
        ];
      }
      const backendChecks = (await response.json()) as ICheckResult[];
      return [
        { name: "Backend connection", status: "pass", message: "Frontend reached the backend API." },
        ...backendChecks,
      ];
    } catch (e) {
      return [
        {
          name: "Backend connection",
          status: "fail",
          message: e instanceof Error ? e.message : String(e),
        },
      ];
    }
  }
}
