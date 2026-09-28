import { Role } from "./IRole";

export interface IAccessEntry {
  email: string;
  role: Role;
}

export interface ISettings {
  msFormUrl: string;
  flowId: string;
  access: IAccessEntry[];
}

export const DEFAULT_SETTINGS: ISettings = {
  msFormUrl: "",
  flowId: "",
  access: [],
};

/** Discard retired connection and email-to-role settings when loading old JSON. */
export function normalizeSettings(value: unknown): ISettings {
  const stored = value && typeof value === "object" ? value as Record<string, unknown> : {};
  return {
    msFormUrl: typeof stored.msFormUrl === "string" ? stored.msFormUrl : "",
    flowId: typeof stored.flowId === "string" ? stored.flowId : "",
    access: [],
  };
}
