package com.timematters.qstar.service;

import java.time.LocalDate;
import java.time.OffsetDateTime;
import java.util.Set;

public final class CalendarDates {
    public static final Set<String> FIELDS =
            Set.of(
                    "reportDate",
                    "dueDate",
                    "implementationDate",
                    "verifiedDate",
                    "closedDate",
                    "holdUntil");

    private CalendarDates() {}

    /**
     * Preserve the calendar component of SharePoint DateOnly envelopes, never convert it through
     * UTC.
     */
    public static String normalize(String value) {
        if (value == null || value.isBlank()) return "";
        if (value.matches("[0-9]{4}-[0-9]{2}-[0-9]{2}")) return LocalDate.parse(value).toString();
        OffsetDateTime timestamp = OffsetDateTime.parse(value);
        return timestamp.toLocalDate().toString();
    }
}
