export type SPUserValue = {
  Id?: number;
  Title?: string;
  EMail?: string;
  LoginName?: string;
};

export interface IPersonValue {
  id?: number;
  displayName: string;
  email: string;
}

export function readPersonValue(raw: SPUserValue | string | undefined): IPersonValue {
  if (!raw) return { displayName: "", email: "" };
  if (typeof raw === "string") return { displayName: raw, email: "" };
  return {
    id: raw.Id,
    displayName: raw.Title || "",
    email: raw.EMail || raw.LoginName || "",
  };
}
