export type Role = "admin" | "qm" | "owner" | "reader";

export interface IResolvedRole {
  role: Role;
  source: "sharepoint" | "development" | "backend";
  matchedGroup?: string;
  user?: IBackendUser;
  connection?: IQstarConnection;
}
import { IBackendUser, IQstarConnection } from "./IConnection";
