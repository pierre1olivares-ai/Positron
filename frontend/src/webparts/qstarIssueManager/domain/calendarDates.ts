/** Calendar dates are YYYY-MM-DD values; event timestamps retain their time zone. */
export type CalendarDateValue = string | Date;

const DAY_MILLISECONDS = 86400000;

function twoDigits(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

function fourDigits(value: number): string {
  return (`0000${value}`).slice(-4);
}

function utcDate(value: CalendarDateValue): Date | undefined {
  const normalized = value instanceof Date ? formatLocalDate(value) : normalizeDateOnly(value);
  if (!normalized) return undefined;
  const parts = normalized.split("-");
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return date;
}

function formatUTCDate(date: Date): string {
  return `${fourDigits(date.getUTCFullYear())}-${twoDigits(date.getUTCMonth() + 1)}-${twoDigits(date.getUTCDate())}`;
}

/** Format a user's local calendar day without converting midnight to UTC. */
export function formatLocalDate(date: Date): string {
  if (isNaN(date.getTime())) return "";
  return `${fourDigits(date.getFullYear())}-${twoDigits(date.getMonth() + 1)}-${twoDigits(date.getDate())}`;
}

export function todayDate(now: Date = new Date()): string {
  return formatLocalDate(now);
}

/**
 * SharePoint returns date-only fields in ISO DateTime envelopes. Preserve their
 * calendar component instead of shifting their day through the browser's zone.
 * Use this only for calendar fields, never event timestamps such as ClosedAt.
 */
// SharePoint represents absent DateOnly fields with JSON null.
// eslint-disable-next-line @rushstack/no-new-null
export function normalizeDateOnly(value: string | undefined | null): string {
  if (!value) return "";
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:T.+)?$/.exec(value);
  if (!match || (value.length > 10 && isNaN(new Date(value).getTime()))) return "";
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return "";
  return `${match[1]}-${match[2]}-${match[3]}`;
}

export function addCalendarDays(value: CalendarDateValue, days: number): string {
  const date = utcDate(value);
  if (!date || !isFinite(days)) return "";
  date.setUTCDate(date.getUTCDate() + days);
  return formatUTCDate(date);
}

/** Match calendar-month arithmetic: December 31 + 2 months is February's last day. */
export function addCalendarMonths(value: CalendarDateValue, months: number): string {
  const date = utcDate(value);
  if (!date || !isFinite(months)) return "";
  const originalDay = date.getUTCDate();
  date.setUTCDate(1);
  date.setUTCMonth(date.getUTCMonth() + months);
  const lastDay = new Date(date.getTime());
  lastDay.setUTCMonth(lastDay.getUTCMonth() + 1, 0);
  date.setUTCDate(Math.min(originalDay, lastDay.getUTCDate()));
  return formatUTCDate(date);
}

/** Signed whole calendar days from a to b; unaffected by daylight-saving changes. */
export function calendarDaysBetween(a: CalendarDateValue, b: CalendarDateValue): number {
  const first = utcDate(a);
  const second = utcDate(b);
  return first && second ? (second.getTime() - first.getTime()) / DAY_MILLISECONDS : NaN;
}

export function sameCalendarDay(a: CalendarDateValue, b: CalendarDateValue): boolean {
  return calendarDaysBetween(a, b) === 0;
}

export function isDateInYearThroughToday(value: CalendarDateValue, asOf: CalendarDateValue = todayDate()): boolean {
  const date = utcDate(value);
  const limit = utcDate(asOf);
  return !!date && !!limit && date.getUTCFullYear() === limit.getUTCFullYear() && date.getTime() <= limit.getTime();
}
