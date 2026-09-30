import { IIssue } from "../models/IIssue";
import { positiveId } from "../services/sharePointValues";

export interface IIssueHistory {
  issueId: number;
  eTag?: string;
  current: Record<string, unknown>;
  versions: unknown[];
  complete: boolean;
}

export function historyKey(issue: IIssue): string {
  return JSON.stringify([issue.id, issue.eTag, issue.triaged, issue.taskCreated, issue.status]);
}

function timestamp(value: unknown): number | undefined {
  if (typeof value !== "string" || !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)$/.test(value)) return undefined;
  const [year, month, day] = value.slice(0, 10).split("-").map(Number);
  if (month < 1 || month > 12 || day < 1 || day > new Date(Date.UTC(year, month, 0)).getUTCDate()) return undefined;
  const time = Date.parse(value);
  return isFinite(time) ? time : undefined;
}

function state(row: Record<string, unknown>): { triaged: boolean; taskCreated: string; status: string; open: boolean } | undefined {
  if (["Yes", "No"].indexOf(row.Triaged as string) < 0 || ["Yes", "No"].indexOf(row.TaskCreated as string) < 0 || !Object.prototype.hasOwnProperty.call(row, "Status")) return undefined;
  const triaged = row.Triaged === "Yes";
  const status = row.Status === null || row.Status === undefined ? "" : row.Status;
  const active = ["Created", "In Progress", "On Hold", "Under Testing/Revision"];
  if (typeof status !== "string" || [...active, "Closed", "Rejected", ...(!triaged ? [""] : [])].indexOf(status) < 0) return undefined;
  return { triaged, taskCreated: row.TaskCreated as string, status, open: triaged && row.TaskCreated === "Yes" && active.indexOf(status) >= 0 };
}

function ordinal(label: unknown): number | undefined {
  if (typeof label !== "string" || !/^[1-9]\d*\.0$/.test(label)) return undefined;
  const major = Number(label.slice(0, -2));
  return positiveId(major) ? major : undefined;
}

/** Read the state immediately before an instant; callers supply local period boundaries as Dates. */
export function issueOpenBefore(issue: IIssue, history: IIssueHistory | undefined, boundary: Date): boolean | undefined {
  if (!history?.complete || history.issueId !== issue.id || !issue.eTag || history.eTag !== issue.eTag || !history.current || !Array.isArray(history.versions)) return undefined;
  const current = history.current;
  const currentState = state(current);
  const created = timestamp(current.Created);
  const modified = timestamp(current.Modified);
  const latestOrdinal = ordinal(current.OData__UIVersionString);
  const before = boundary.getTime();
  if (current.Id !== issue.id || !currentState || created === undefined || modified === undefined || modified < created || !latestOrdinal || !isFinite(before) ||
      currentState.triaged !== issue.triaged || currentState.taskCreated !== issue.taskCreated || currentState.status !== (issue.status || "")) return undefined;
  if (before <= created) return false;
  const rows = history.versions as Record<string, unknown>[];
  if (!rows.length || rows.some(row => !row || !ordinal(row.VersionLabel) || !positiveId(row.VersionId) || typeof row.IsCurrentVersion !== "boolean")) return undefined;
  const ordered = [...rows].sort((a, b) => ordinal(a.VersionLabel)! - ordinal(b.VersionLabel)!);
  if (new Set(ordered.map(row => row.VersionLabel)).size !== rows.length || new Set(ordered.map(row => row.VersionId)).size !== rows.length) return undefined;
  const latest = ordered[ordered.length - 1];
  const latestState = state(latest);
  if (latest.VersionLabel !== current.OData__UIVersionString || !latest.IsCurrentVersion || ordered.slice(0, -1).some(row => row.IsCurrentVersion) ||
      timestamp(latest.Created) !== modified || !latestState || latestState.triaged !== currentState.triaged || latestState.taskCreated !== currentState.taskCreated || latestState.status !== currentState.status) return undefined;
  let nextOrdinal = latestOrdinal + 1;
  let nextTime = Infinity;
  let selected: boolean | undefined;
  for (let index = ordered.length - 1; index >= 0; index--) {
    const row = ordered[index];
    const time = timestamp(row.Created);
    const value = state(row);
    const version = ordinal(row.VersionLabel)!;
    if (version !== nextOrdinal - 1 || time === undefined || time < created || time > nextTime || !value) break;
    if (time < before && selected === undefined) selected = value.open;
    nextOrdinal = version;
    nextTime = time;
  }
  return selected;
}

export function historicalBacklog(issues: IIssue[], histories: Record<number, IIssueHistory | undefined>, boundary: Date): number | undefined {
  let total = 0;
  for (const issue of issues) {
    const open = issueOpenBefore(issue, histories[issue.id], boundary);
    if (open === undefined) return undefined;
    if (open) total++;
  }
  return total;
}
