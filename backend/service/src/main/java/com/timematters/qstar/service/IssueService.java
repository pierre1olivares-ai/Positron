package com.timematters.qstar.service;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.api.model.IssueATO;
import com.timematters.qstar.api.model.IssueCreateATO;
import com.timematters.qstar.api.model.IssuePatchATO;
import com.timematters.qstar.api.model.ProgressCreateATO;
import com.timematters.qstar.api.model.ProgressLogEntryATO;
import com.timematters.qstar.configuration.security.CallerContext;
import com.timematters.qstar.configuration.security.CurrentUserProvider;
import com.timematters.qstar.infrastructure.sharepoint.IssueRepository;
import com.timematters.qstar.infrastructure.sharepoint.SharePointUser;
import com.timematters.qstar.model.Issue;
import com.timematters.qstar.model.mapper.IssueMapper;
import com.timematters.qstar.model.mapper.ProgressLogEntryMapper;
import java.time.Clock;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.stereotype.Service;
import org.springframework.web.client.HttpStatusCodeException;

@Service
public class IssueService {
    private static final Map<String, Integer> DUE_DAYS =
            Map.of("Critical", 7, "High", 14, "Medium", 30, "Low", 60);
    private static final Set<String> OWNER_FIELDS =
            Set.of("status", "implementationDate", "holdReason", "holdUntil");
    private static final Set<String> OWNER_STATUS =
            Set.of("Created", "In Progress", "On Hold", IssueLifecycle.TEST);
    private final IssueRepository repository;
    private final CurrentUserProvider users;
    private final ObjectMapper json;
    private final Clock clock;
    private final IssueMapper issues = new IssueMapper();
    private final ProgressLogEntryMapper progress = new ProgressLogEntryMapper();

    public IssueService(
            IssueRepository repository, CurrentUserProvider users, ObjectMapper json, Clock clock) {
        this.repository = repository;
        this.users = users;
        this.json = json;
        this.clock = clock;
    }

    public List<IssueATO> getAllIssues() {
        users.currentUser();
        return repository.findAll().stream().map(issues::toATO).toList();
    }

    public IssueATO getIssue(long id) {
        users.currentUser();
        return issues.toATO(repository.findById(id));
    }

    public IssueATO createIssue(IssueCreateATO create) {
        CallerContext caller = users.currentUser();
        manager(caller);
        Map<String, Object> input = json.convertValue(create, new TypeReference<>() {});
        input.entrySet().removeIf(entry -> entry.getValue() == null);
        requireText(input, "shortSummary");
        requireText(input, "description");
        if (!DUE_DAYS.containsKey(text(input.get("severity")))) invalid("Choose a valid severity.");
        input.put(
                "reportDate",
                input.containsKey("reportDate")
                        ? inputDate(input.get("reportDate"), "reportDate")
                        : LocalDate.now(clock).toString());
        requireText(input, "reportDate");
        input.put("triaged", false);
        input.put("taskCreated", "No");
        input.put("reminderCycle", "initial");
        input.putIfAbsent("alreadyInContact", "No");
        if (!input.containsKey("createdById") && !input.containsKey("createdByEmail")) {
            input.put("createdBy", caller.displayName());
            input.put("createdById", caller.sharePointUserId());
            input.put("createdByEmail", caller.email());
        }
        return issues.toATO(repository.create(input));
    }

    public Object updateIssue(long id, IssuePatchATO requested, String originalETag) {
        if (originalETag == null || originalETag.isBlank() || originalETag.equals("*")) {
            throw new IssueOperationException(
                    428,
                    "ISSUE_VERSION_REQUIRED",
                    "Reload the issue and supply its original ETag before saving.");
        }
        CallerContext caller = users.currentUser();
        Issue current = repository.findById(id);
        Map<String, Object> patch = json.convertValue(requested, new TypeReference<>() {});
        // JsonNullable serializes only explicitly supplied properties, including explicit nulls.
        authorizePatch(caller, current, patch);
        if (!Objects.equals(current.getETag(), originalETag)) throw conflict(id, current);
        validate(patch);
        IssueLifecycle.apply(current, patch, clock);
        if (Boolean.TRUE.equals(patch.get("triaged"))
                && !Boolean.TRUE.equals(current.getTriaged())
                && !patch.containsKey("dueDate")
                && text(current.getDueDate()).isEmpty()) {
            patch.put(
                    "dueDate",
                    LocalDate.parse(CalendarDates.normalize(current.getReportDate()))
                            .plusDays(DUE_DAYS.getOrDefault(current.getSeverity(), 30))
                            .toString());
        }
        if ("owner".equals(caller.role())
                && patch.containsKey("status")
                && !Objects.equals(current.getStatus(), patch.get("status"))) {
            patch.put("ownerUpdate", true);
            patch.put("ownerUpdateAt", clock.instant().toString());
            patch.put(
                    "ownerUpdateText",
                    "Status changed from \""
                            + current.getStatus()
                            + "\" to \""
                            + patch.get("status")
                            + "\".");
        }
        if (patch.isEmpty()) return issues.toATO(current);
        try {
            repository.update(id, patch, originalETag);
        } catch (HttpStatusCodeException e) {
            if (e.getStatusCode().value() != 412) throw e;
            Issue latest = null;
            try {
                latest = repository.findById(id);
            } catch (Exception ignored) {
            }
            throw conflict(id, latest);
        }
        try {
            return issues.toATO(repository.findById(id));
        } catch (Exception e) {
            return Map.of(
                    "saved",
                    true,
                    "issueId",
                    id,
                    "saveWarning",
                    "Your changes were saved. Reload this issue before editing again; do not repeat the save.");
        }
    }

    public ProgressLogEntryATO addProgressLogEntry(long id, ProgressCreateATO input) {
        CallerContext caller = users.currentUser();
        Issue issue = repository.findById(id);
        contributor(caller, issue);
        if (Set.of("Closed", "Rejected").contains(text(issue.getStatus())))
            invalid("Closed or rejected issues cannot receive progress updates.");
        String text = input.getText();
        if (text == null || text.isBlank() || text.length() > 63000)
            invalid("Enter a progress update of 1 to 63000 characters.");
        return progress.toATO(
                repository.append(
                        id,
                        text,
                        new SharePointUser(
                                caller.sharePointUserId(), caller.displayName(), caller.email())));
    }

    public ProgressLogEntryATO getProgressEntry(long id, long entryId) {
        users.currentUser();
        repository.findById(id);
        return progress.toATO(repository.progressEntry(id, entryId));
    }

    private void authorizePatch(CallerContext caller, Issue issue, Map<String, Object> patch) {
        contributor(caller, issue);
        if (!"owner".equals(caller.role())) return;
        // Older beta clients send these derived fields; ignore them and derive trusted values
        // below.
        patch.remove("ownerUpdate");
        patch.remove("ownerUpdateAt");
        patch.remove("ownerUpdateText");
        if (IssueLifecycle.TEST.equals(issue.getStatus()))
            throw new AccessDeniedException(
                    "Only the Quality Team can change an issue under effectiveness testing.");
        if (!OWNER_FIELDS.containsAll(patch.keySet()))
            throw new AccessDeniedException(
                    "Task owners can update only their task progress, implementation date and hold details.");
        if (patch.containsKey("status") && !OWNER_STATUS.contains(text(patch.get("status"))))
            throw new AccessDeniedException(
                    "Only the Quality Team can close, reject or reopen issues.");
    }

    private void contributor(CallerContext caller, Issue issue) {
        if (caller.role().equals("admin") || caller.role().equals("qm")) return;
        if (!caller.role().equals("owner")
                || !Objects.equals(issue.getTaskOwnerId(), caller.sharePointUserId())
                || Set.of("Closed", "Rejected").contains(text(issue.getStatus()))) {
            throw new AccessDeniedException(
                    "Only the current task owner or the Quality Team can change this issue.");
        }
    }

    private void manager(CallerContext caller) {
        if (!Set.of("admin", "qm").contains(caller.role()))
            throw new AccessDeniedException("Only the Quality Team can create an issue.");
    }

    private IssueOperationException conflict(long id, Issue latest) {
        var meta = new LinkedHashMap<String, Object>();
        meta.put("issueId", id);
        if (latest != null) meta.put("freshIssue", issues.toATO(latest));
        return new IssueOperationException(
                412,
                "ISSUE_CONFLICT",
                "This issue changed while you were editing. Keep your draft and reload explicitly.",
                meta);
    }

    private void validate(Map<String, Object> patch) {
        for (String name : CalendarDates.FIELDS)
            if (patch.containsKey(name)) patch.put(name, inputDate(patch.get(name), name));
        if (patch.containsKey("severity") && !DUE_DAYS.containsKey(text(patch.get("severity"))))
            invalid("Choose a valid severity.");
        for (String name : Set.of("taskCreated", "alreadyInContact"))
            if (patch.containsKey(name) && !Set.of("Yes", "No").contains(text(patch.get(name))))
                invalid("Choose Yes or No for " + name + ".");
        for (String name : Set.of("triaged", "ownerUpdate"))
            if (patch.containsKey(name) && !(patch.get(name) instanceof Boolean))
                invalid(name + " must be true or false.");
        for (String name : Set.of("shortSummary", "description"))
            if (patch.containsKey(name)) requireText(patch, name);
    }

    private static void requireText(Map<String, Object> values, String name) {
        if (text(values.get(name)).isBlank()) invalid(name + " is required.");
    }

    private static String inputDate(Object value, String field) {
        try {
            return CalendarDates.normalize(text(value));
        } catch (java.time.DateTimeException invalidDate) {
            throw new IssueOperationException(
                    400, "INVALID_ISSUE", "Enter a valid calendar date for " + field + ".");
        }
    }

    private static String text(Object value) {
        return value == null ? "" : value.toString();
    }

    private static void invalid(String message) {
        throw new IssueOperationException(400, "INVALID_ISSUE", message);
    }
}
