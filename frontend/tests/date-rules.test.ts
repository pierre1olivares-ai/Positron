import assert from "node:assert/strict";
import test from "node:test";

import {
  addCalendarDays,
  addCalendarMonths,
  calendarDaysBetween,
  formatLocalDate,
  isDateInYearThroughToday,
  normalizeDateOnly,
  sameCalendarDay,
  todayDate,
} from "../src/webparts/qstarIssueManager/domain/calendarDates";

test("local calendar dates do not convert local midnight through UTC", () => {
  const localMidnight = new Date(2026, 8, 28, 0, 0, 0);
  assert.equal(formatLocalDate(localMidnight), "2026-09-28");
  assert.equal(todayDate(localMidnight), "2026-09-28");
  assert.equal(addCalendarDays(localMidnight, 7), "2026-10-05");
});

test("SharePoint date-only envelopes become valid native date input values", () => {
  assert.equal(normalizeDateOnly("2026-10-01T00:00:00Z"), "2026-10-01");
  assert.equal(normalizeDateOnly("2026-10-01T00:00:00.000Z"), "2026-10-01");
  assert.equal(normalizeDateOnly("2026-10-01"), "2026-10-01");
  assert.equal(normalizeDateOnly(undefined), "");
  assert.equal(normalizeDateOnly("2026-02-29"), "");
  assert.equal(normalizeDateOnly("2026-02-29T00:00:00Z"), "");
  assert.equal(normalizeDateOnly("2024-02-29T00:00:00Z"), "2024-02-29");
  assert.equal(normalizeDateOnly("2026-10-01Tnot-a-time"), "");
});

test("calendar day arithmetic survives daylight-saving and year boundaries", () => {
  assert.equal(addCalendarDays("2026-03-28", 2), "2026-03-30");
  assert.equal(calendarDaysBetween("2026-03-28", "2026-03-30"), 2);
  assert.equal(calendarDaysBetween("2026-10-24", "2026-10-26"), 2);
  assert.equal(addCalendarDays("2026-12-31", 1), "2027-01-01");
  assert.equal(addCalendarDays("2026-01-01", -1), "2025-12-31");
});

test("month arithmetic clamps to the final day of the destination month", () => {
  assert.equal(addCalendarMonths("2026-12-31", 2), "2027-02-28");
  assert.equal(addCalendarMonths("2023-12-31", 2), "2024-02-29");
  assert.equal(addCalendarMonths("2026-07-31", 2), "2026-09-30");
  assert.equal(addCalendarMonths("2026-01-31", 2), "2026-03-31");
});

test("today's timestamp reports are included in year-to-date metrics", () => {
  const asOf = new Date(2026, 8, 28, 12, 0, 0);
  assert.equal(isDateInYearThroughToday("2026-09-28T10:57:42Z", asOf), true);
  assert.equal(isDateInYearThroughToday("2026-09-28", asOf), true);
  assert.equal(isDateInYearThroughToday("2026-09-29", asOf), false);
  assert.equal(isDateInYearThroughToday("2025-12-31", asOf), false);
  assert.equal(isDateInYearThroughToday("", asOf), false);
});

test("a timestamp envelope and a date-only value match the same reminder day", () => {
  assert.equal(sameCalendarDay("2026-09-28T00:00:00Z", "2026-09-28"), true);
  assert.equal(sameCalendarDay("2026-09-27T00:00:00Z", "2026-09-28"), false);
  assert.equal(sameCalendarDay("", ""), false);
});
