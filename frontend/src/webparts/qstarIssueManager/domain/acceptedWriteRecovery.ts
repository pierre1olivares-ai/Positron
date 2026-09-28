export interface IAcceptedReceipt {
  message: string;
  submitted: string;
}
export type AcceptedReceipts = Record<string, IAcceptedReceipt>;

const memory = new Map<string, AcceptedReceipts>();
const listeners = new Map<string, Set<(receipts: AcceptedReceipts) => void>>();

/** Connection/user-scoped receipts survive component remounts and same-tab reloads. */
export function readAcceptedReceipts(key: string): AcceptedReceipts {
  const current = memory.get(key);
  if (current) return current;
  try {
    const stored = window.sessionStorage.getItem(key);
    if (stored) {
      const parsed = JSON.parse(stored) as AcceptedReceipts;
      const valid: AcceptedReceipts = {};
      Object.keys(parsed).forEach((operation) => {
        const receipt = parsed[operation];
        if ((operation === "create" || /^progress:[1-9]\d*$/.test(operation)) && receipt &&
            typeof receipt.message === "string" && typeof receipt.submitted === "string") valid[operation] = receipt;
      });
      memory.set(key, valid);
      return valid;
    }
  } catch { /* The in-memory guard still survives remounts if browser storage is unavailable. */ }
  return memory.get(key) || {};
}

export function writeAcceptedReceipts(key: string, receipts: AcceptedReceipts): void {
  memory.set(key, receipts);
  try {
    if (Object.keys(receipts).length) window.sessionStorage.setItem(key, JSON.stringify(receipts));
    else window.sessionStorage.removeItem(key);
  } catch { /* Do not turn an accepted remote write into a retryable storage failure. */ }
  listeners.get(key)?.forEach((listener) => listener(receipts));
}

/** An older instance can finish a request after the same connection has remounted. */
export function subscribeAcceptedReceipts(key: string, listener: (receipts: AcceptedReceipts) => void): () => void {
  const subscribers = listeners.get(key) || new Set<(receipts: AcceptedReceipts) => void>();
  subscribers.add(listener);
  listeners.set(key, subscribers);
  return () => {
    subscribers.delete(listener);
    if (!subscribers.size) listeners.delete(key);
  };
}
