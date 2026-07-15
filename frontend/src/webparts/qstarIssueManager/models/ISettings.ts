import { Role } from "./IRole";

export interface IAccessEntry {
  email: string;
  role: Role;
}

export interface ISettings {
  msFormUrl: string;
  flowId: string;
  spSiteUrl: string;
  spListName: string;
  connected: boolean;
  lastTested: string;
  access: IAccessEntry[];
}

export const DEFAULT_SETTINGS: ISettings = {
  msFormUrl: "",
  flowId: "",
  spSiteUrl: "",
  spListName: "",
  connected: false,
  lastTested: "",
  access: [],
};
