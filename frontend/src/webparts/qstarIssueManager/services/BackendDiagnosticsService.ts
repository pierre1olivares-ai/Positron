import { ICheckResult } from "./ConnectionDiagnosticsService";
import { backendJson, IBackendTransport } from "./BackendApiClient";

export class BackendDiagnosticsService {
  constructor(private client: IBackendTransport) {}

  public async run(): Promise<ICheckResult[]> {
    try {
      const checks = await backendJson(this.client, "/diagnostics");
      if (!Array.isArray(checks) || checks.some((check: ICheckResult) => !check || typeof check.name !== "string" ||
        typeof check.message !== "string" || ["pass", "warn", "fail"].indexOf(check.status) < 0)) throw new Error("The backend returned invalid diagnostics.");
      return [{ name: "Backend connection", status: "pass", message: "Frontend reached the backend API." }, ...checks];
    } catch (error) {
      return [{ name: "Backend connection", status: "fail", message: error instanceof Error ? error.message : String(error) }];
    }
  }
}
