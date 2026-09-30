import { IResolvedRole } from "../models/IRole";
import { IRoleResolver } from "./IRoleResolver";
import { backendJson, IBackendTransport, parseBackendSession } from "./BackendApiClient";

export class BackendRoleResolver implements IRoleResolver {
  constructor(private client: IBackendTransport, private baseUrl: string) {}

  public async resolve(): Promise<IResolvedRole> {
    return parseBackendSession(await backendJson(this.client, "/me"), this.baseUrl);
  }
}
