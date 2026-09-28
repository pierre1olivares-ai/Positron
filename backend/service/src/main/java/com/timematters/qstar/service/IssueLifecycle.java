package com.timematters.qstar.service;

import com.timematters.qstar.model.Issue;
import java.time.Clock;
import java.time.LocalDate;
import java.util.Map;
import java.util.Set;

/** Server equivalent of the direct SharePoint client's transition boundary. */
public final class IssueLifecycle {
    public static final String TEST = "Under Testing/Revision";
    private static final Set<String> STATUSES =
            Set.of("Created", "In Progress", TEST, "On Hold", "Closed", "Rejected");

    private IssueLifecycle() {}

    public static void apply(Issue current, Map<String, Object> patch, Clock clock) {
        LocalDate today = LocalDate.now(clock);
        for (String field : CalendarDates.FIELDS)
            if (patch.containsKey(field)) {
                patch.put(field, CalendarDates.normalize(text(patch.get(field))));
            }
        String status = value(patch, "status", current.getStatus());
        if (patch.containsKey("status") && status.isBlank())
            invalid("Choose a valid issue status.");
        if (!status.isEmpty() && !STATUSES.contains(status))
            invalid("Choose a valid issue status.");
        boolean changed = !status.equals(text(current.getStatus()));
        String category = value(patch, "transformedInto", current.getTransformedInto());
        boolean nc = category.equals("NC Minor") || category.equals("NC Major");
        if ("Closed".equals(current.getStatus())) {
            if (!changed || !status.equals("In Progress"))
                invalid("Re-open the closed issue before editing it.");
            for (String field :
                    Set.of(
                            "closedDate",
                            "closedAt",
                            "verifiedBy",
                            "verifiedByEmail",
                            "verifiedDate",
                            "implementationDate",
                            "effectivenessCheck")) patch.put(field, "");
            patch.put("verifiedById", null);
            patch.put("reminderCycle", clock.instant().toString());
            return;
        }
        if (status.equals(TEST)) {
            if (!nc) invalid("Only a nonconformity can enter effectiveness testing.");
            String date = value(patch, "implementationDate", current.getImplementationDate());
            LocalDate implementation = date.isBlank() ? today : LocalDate.parse(date);
            if (implementation.isAfter(today)) invalid("Implementation must be today or earlier.");
            if (changed
                    || patch.containsKey("implementationDate")
                    || text(current.getImplementationDate()).isEmpty())
                patch.put("implementationDate", implementation.toString());
        }
        if (status.equals("On Hold")) {
            if (value(patch, "holdReason", current.getHoldReason()).isBlank())
                invalid("Enter a reason before putting the issue on hold.");
            String date = value(patch, "holdUntil", current.getHoldUntil());
            if (date.isBlank()) invalid("Enter a resume date before putting the issue on hold.");
            if ((changed || patch.containsKey("holdUntil"))
                    && LocalDate.parse(date).isBefore(today))
                invalid("The resume date must be today or later.");
        }
        if (status.equals("Closed") && changed) {
            if (nc) {
                if (!TEST.equals(current.getStatus()))
                    invalid(
                            "Start the 2-month effectiveness test before closing this nonconformity.");
                String date = value(patch, "implementationDate", current.getImplementationDate());
                if (date.isBlank() || LocalDate.parse(date).plusMonths(2).isAfter(today))
                    invalid(
                            "Complete the 2-month effectiveness test before closing this nonconformity.");
                String verifier = value(patch, "verifiedBy", current.getVerifiedBy());
                String email = value(patch, "verifiedByEmail", current.getVerifiedByEmail());
                Object id =
                        patch.containsKey("verifiedById")
                                ? patch.get("verifiedById")
                                : current.getVerifiedById();
                if (verifier.isBlank()
                        || (email.isBlank()
                                && (!(id instanceof Number number) || number.longValue() <= 0)))
                    invalid(
                            "Record the verifier's name and Microsoft 365 identity before closing.");
                String verified = value(patch, "verifiedDate", current.getVerifiedDate());
                patch.put("verifiedDate", verified.isBlank() ? today.toString() : verified);
            }
            patch.put("closedDate", today.toString());
            patch.put("closedAt", clock.instant().toString());
        }
        if (status.equals("Rejected")) patch.put("taskCreated", "No");
    }

    public static String text(Object value) {
        return value == null ? "" : value.toString();
    }

    private static String value(Map<String, Object> patch, String name, String previous) {
        return patch.containsKey(name) ? text(patch.get(name)) : text(previous);
    }

    private static void invalid(String message) {
        throw new IssueOperationException(400, "INVALID_TRANSITION", message);
    }
}
