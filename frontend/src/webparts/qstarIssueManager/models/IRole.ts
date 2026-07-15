export type Role = "admin" | "qm" | "owner" | "reader";

export interface IResolvedRole {
  role: Role;
  source: "sharepoint" | "development";
  matchedGroup?: string;
}
