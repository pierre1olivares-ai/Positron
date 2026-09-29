import { useEffect, useRef, useState } from "react";
import { IIssue } from "../models/IIssue";
import { IDataService } from "../services/IDataService";
import { historyKey, IIssueHistory } from "../domain/issueHistory";

export function useIssueHistory(issues: IIssue[], service?: IDataService): Record<number, IIssueHistory | undefined> {
  const cache = useRef(new Map<string, IIssueHistory | undefined>());
  const workers = useRef({ active: 0, wake: (): void => {} });
  const previousService = useRef(service);
  if (previousService.current !== service) {
    cache.current.clear();
    previousService.current = service;
  }
  const [, refresh] = useState(0);
  const signature = JSON.stringify(issues.map(historyKey));
  useEffect(() => {
    let cancelled = false;
    let cursor = 0;
    const pending = issues.filter(issue => issue.eTag && !cache.current.has(historyKey(issue)));
    const schedule = (): void => {
      while (workers.current.active < 4 && cursor < pending.length) {
        if (cancelled) return;
        const issue = pending[cursor++];
        workers.current.active++;
        const finish = (history: IIssueHistory | undefined): void => {
          workers.current.active--;
          if (!cancelled) {
            cache.current.set(historyKey(issue), history);
            refresh(value => value + 1);
          }
          workers.current.wake();
        };
        Promise.resolve().then(() => service?.getIssueHistory(issue.id)).then(finish, () => finish(undefined));
      }
    };
    const activeKeys = new Set(issues.map(historyKey));
    cache.current.forEach((_, key) => { if (!activeKeys.has(key)) cache.current.delete(key); });
    workers.current.wake = schedule;
    schedule();
    return () => { cancelled = true; workers.current.wake = (): void => {}; };
  }, [service, signature]);
  const result: Record<number, IIssueHistory | undefined> = {};
  for (const issue of issues) result[issue.id] = cache.current.get(historyKey(issue));
  return result;
}
