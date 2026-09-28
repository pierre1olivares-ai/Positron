export type SPUserValue = {
  Id?: number;
  Title?: string;
  EMail?: string;
  Name?: string;
};

export interface IPersonValue {
  id?: number;
  displayName: string;
  email: string;
}

export function positiveId(value: unknown): value is number {
  return typeof value === "number" && isFinite(value) && value > 0 && value <= 9007199254740991 && Math.floor(value) === value;
}

export function readPersonValue(raw: SPUserValue | string | undefined): IPersonValue {
  if (!raw) return { displayName: "", email: "" };
  if (typeof raw === "string") return { displayName: raw, email: "" };
  return {
    id: raw.Id,
    displayName: raw.Title || "",
    email: raw.EMail || (raw.Name || "").split("|").pop() || "",
  };
}
