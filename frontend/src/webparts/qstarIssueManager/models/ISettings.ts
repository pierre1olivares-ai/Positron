export type Role = "admin" | "qm" | "owner" | "reader";

export interface IAccessEntry {
  email: string;
  role: Role;
}

export interface ISettings {
  msFormUrl: string;
  flowId: string;
  spSiteUrl: string;
  spListName: string;
  tenantId: string;
  clientId: string;
  connected: boolean;
  lastTested: string;
  access: IAccessEntry[];
}

export const DEFAULT_SETTINGS: ISettings = {
  msFormUrl: "",
  flowId: "",
  spSiteUrl: "",
  spListName: "",
  tenantId: "",
  clientId: "",
  connected: false,
  lastTested: "",
  access: [],
};
