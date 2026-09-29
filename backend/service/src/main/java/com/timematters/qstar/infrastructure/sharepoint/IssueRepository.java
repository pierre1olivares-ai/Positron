package com.timematters.qstar.infrastructure.sharepoint;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.model.Issue;
import com.timematters.qstar.model.ProgressLogEntry;
import com.timematters.qstar.service.CalendarDates;
import com.timematters.qstar.service.IssueOperationException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.springframework.stereotype.Repository;

@Repository
public class IssueRepository {
    public static final Map<String, String> FIELDS =
            Map.ofEntries(
                    Map.entry("qsNumber", "QsNumber"), Map.entry("triaged", "Triaged"),
                    Map.entry("status", "Status"), Map.entry("taskCreated", "TaskCreated"),
                    Map.entry("transformedInto", "TransformedInto"),
                            Map.entry("shortSummary", "ShortSummary"),
                    Map.entry("description", "Description"),
                            Map.entry("immediateAction", "ImmediateAction"),
                    Map.entry("severity", "Severity"), Map.entry("reportDate", "ReportDate"),
                    Map.entry("departmentBU", "DepartmentBU"), Map.entry("region", "Region"),
                    Map.entry("alreadyInContact", "AlreadyInContact"),
                            Map.entry("deviationType", "DeviationType"),
                    Map.entry("issueOrigin", "Origin"),
                            Map.entry("additionalComments", "AdditionalComments"),
                    Map.entry("followUp", "FollowUp"), Map.entry("ownerBU", "EscalationBU"),
                    Map.entry("dueDate", "DueDate"), Map.entry("rootCause", "RootCause"),
                    Map.entry("correctiveAction", "CorrectiveAction"),
                            Map.entry("implementationDate", "ImplementationDate"),
                    Map.entry("effectivenessCheck", "EffectivenessCheck"),
                            Map.entry("verifiedDate", "VerifiedDate"),
                    Map.entry("closedDate", "ClosedDate"), Map.entry("closedAt", "ClosedAt"),
                    Map.entry("holdReason", "HoldReason"), Map.entry("holdUntil", "HoldUntil"),
                    Map.entry("ownerUpdate", "OwnerUpdate"),
                            Map.entry("ownerUpdateAt", "OwnerUpdateAt"),
                    Map.entry("ownerUpdateText", "OwnerUpdateText"),
                            Map.entry("reminderCycle", "ReminderCycle"));
    public static final Map<String, String> PEOPLE =
            Map.of("createdBy", "ReportedBy", "taskOwner", "TaskOwner", "verifiedBy", "VerifiedBy");
    private static final Set<String> BOOLS = Set.of("triaged", "ownerUpdate");
    private static final Set<String> TIMESTAMPS = Set.of("closedAt", "ownerUpdateAt");
    private static final String ISSUE_QUERY =
            "$select=*,ReportedBy/Id,ReportedBy/Title,ReportedBy/EMail,ReportedBy/Name,TaskOwner/Id,TaskOwner/Title,TaskOwner/EMail,TaskOwner/Name,VerifiedBy/Id,VerifiedBy/Title,VerifiedBy/EMail,VerifiedBy/Name&$expand=ReportedBy,TaskOwner,VerifiedBy";
    private static final String PROGRESS_QUERY =
            "$select=Id,EntryText,Created,FileDirRef,FSObjType,Author/Id,Author/Title,Author/EMail,Author/Name&$expand=Author";
    private final SharePointRestClient client;
    private final SharePointProperties properties;
    private final ObjectMapper json = new ObjectMapper();

    public IssueRepository(SharePointRestClient client, SharePointProperties properties) {
        this.client = client;
        this.properties = properties;
    }

    public List<Issue> findAll() {
        long offset = referenceOffset();
        String root = progressRoot();
        var logs = new LinkedHashMap<Long, List<ProgressLogEntry>>();
        for (var row :
                client.getItems(
                        properties.getProgressListName(),
                        PROGRESS_QUERY + "&$filter=FSObjType eq 0")) {
            Long parent = parentId(root, row);
            if (parent != null)
                logs.computeIfAbsent(parent, ignored -> new ArrayList<>()).add(toProgress(row));
        }
        return client.getItems(properties.getIssuesListName(), ISSUE_QUERY).stream()
                .map(
                        row ->
                                toIssue(
                                        row,
                                        offset,
                                        logs.getOrDefault(number(row.get("Id")), List.of())))
                .toList();
    }

    public Issue findById(long id) {
        long offset = referenceOffset();
        String folder = folder(id);
        var row = client.getItem(properties.getIssuesListName(), id, ISSUE_QUERY);
        var entries =
                client
                        .getItems(
                                properties.getProgressListName(),
                                PROGRESS_QUERY
                                        + "&$filter=FSObjType eq 0 and FileDirRef eq '"
                                        + SharePointRestClient.literal(folder)
                                        + "'")
                        .stream()
                        .filter(entry -> folder.equals(entry.get("FileDirRef")))
                        .map(this::toProgress)
                        .toList();
        return toIssue(row, offset, entries);
    }

    public Issue create(Map<String, Object> input) {
        long offset = referenceOffset();
        var fields = toFields(input);
        fields.remove("QsNumber");
        Issue fallback = json.convertValue(input, Issue.class);
        fallback.setProgressLog(List.of());
        Map<String, Object> created;
        try {
            created = client.createItem(properties.getIssuesListName(), fields);
        } catch (AcceptedWriteException accepted) {
            fallback.setSaveWarning(accepted.getMessage());
            return fallback;
        }
        Long id;
        try {
            id = created == null ? null : number(created.get("Id"));
        } catch (RuntimeException malformedIdentity) {
            id = null;
        }
        if (id == null || id <= 0) {
            fallback.setSaveWarning(
                    "The report was accepted. Reload before submitting again; its identity was unavailable.");
            return fallback;
        }
        fallback.setId(id);
        long reference;
        try {
            reference = reference(offset, id);
        } catch (RuntimeException unsupportedReference) {
            fallback.setSaveWarning(
                    "The report was saved. Its reference exceeds the supported range; reload and contact an administrator. Do not submit it again.");
            return fallback;
        }
        fallback.setQsNumber(reference);
        fallback.setETag(null);
        String warning = null;
        try {
            String version = eTag(created);
            if (version == null)
                version = eTag(client.getItem(properties.getIssuesListName(), id, "$select=Id"));
            client.updateItem(
                    properties.getIssuesListName(), id, Map.of("QsNumber", reference), version);
        } catch (Exception e) {
            warning =
                    "The report was saved. Its reference is reserved; SharePoint synchronization is pending.";
        }
        try {
            Issue saved = findById(id);
            saved.setSaveWarning(warning);
            return saved;
        } catch (Exception e) {
            fallback.setSaveWarning(
                    "The report was saved. Reload its latest details; do not submit it again.");
            return fallback;
        }
    }

    public void update(long id, Map<String, Object> patch, String eTag) {
        var fields = toFields(patch);
        fields.remove("QsNumber");
        client.updateItem(properties.getIssuesListName(), id, fields, eTag);
    }

    public ProgressLogEntry append(long issueId, String text, SharePointUser caller) {
        String folder = folder(issueId);
        try {
            client.ensureFolder(folder);
        } catch (org.springframework.web.client.HttpClientErrorException e) {
            if (e.getStatusCode().value() != 403 && e.getStatusCode().value() != 404) throw e;
            throw new IssueOperationException(
                    409,
                    "PROGRESS_ACCESS_PENDING",
                    "Progress access is still being prepared. Keep your text and retry after the assignment flow completes.");
        }
        Long id;
        try {
            id = client.appendInFolder(properties.getProgressListName(), folder, text);
        } catch (AcceptedWriteException accepted) {
            id = null;
        }
        if (id != null) {
            try {
                return progressEntry(issueId, id);
            } catch (Exception e) {
                /* The append succeeded. Never repeat it because readback failed. */
            }
        }
        var fallback = new ProgressLogEntry();
        fallback.setId(id);
        fallback.setText(text);
        fallback.setAuthor(caller.displayName());
        fallback.setAuthorId(caller.id());
        fallback.setAuthorEmail(caller.email());
        fallback.setTs("");
        fallback.setSaveWarning(
                id == null
                        ? "Your update may have been posted. Reload the progress log and check before resubmitting."
                        : "Your update was posted. Reload its SharePoint-recorded time and details; do not post it again.");
        return fallback;
    }

    public ProgressLogEntry progressEntry(long issueId, long entryId) {
        var row = client.getItem(properties.getProgressListName(), entryId, PROGRESS_QUERY);
        if (!folder(issueId).equals(row.get("FileDirRef"))
                || !Long.valueOf(0).equals(number(row.get("FSObjType")))) {
            throw new IssueOperationException(
                    404, "PROGRESS_NOT_FOUND", "This progress entry does not belong to the issue.");
        }
        return toProgress(row);
    }

    public long referenceOffset() {
        var rows = client.getItems(properties.getConfigListName(), "$select=Id,ReferenceOffset");
        if (rows.size() != 1)
            throw new IllegalStateException(
                    "Reference allocation requires exactly one provisioned Config item.");
        Long value = number(rows.getFirst().get("ReferenceOffset"));
        if (value == null || value < 1000 || value > 9007199254740991L) {
            throw new IllegalStateException(
                    "ReferenceOffset is invalid. Run the approved provisioner before intake.");
        }
        return value;
    }

    private long reference(long offset, long id) {
        long value = Math.addExact(offset, id);
        if (value > 9007199254740991L)
            throw new IllegalStateException("Reference exceeds the supported integer range.");
        return value;
    }

    private String progressRoot() {
        return client.listInfo(properties.getProgressListName()).rootPath().replaceAll("/$", "");
    }

    private String folder(long id) {
        if (id <= 0) throw new IllegalArgumentException("Invalid issue ID.");
        return progressRoot() + "/issue-" + id;
    }

    private Long parentId(String root, Map<String, Object> entry) {
        String path = text(entry.get("FileDirRef"));
        String prefix = root + "/issue-";
        if (!path.startsWith(prefix)) return null;
        String suffix = path.substring(prefix.length());
        if (!suffix.matches("[1-9][0-9]*")) return null;
        try {
            return Long.valueOf(suffix);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    private Issue toIssue(Map<String, Object> row, long offset, List<ProgressLogEntry> entries) {
        var values = new LinkedHashMap<String, Object>();
        for (var field : FIELDS.entrySet()) {
            String name = field.getKey();
            Object value = row.get(field.getValue());
            if (BOOLS.contains(name)) value = "Yes".equals(value);
            else if (CalendarDates.FIELDS.contains(name))
                value = CalendarDates.normalize(text(value));
            else if (!name.equals("qsNumber")) value = text(value);
            values.put(name, value);
        }
        long id = number(row.get("Id"));
        Long existing = number(row.get("QsNumber"));
        values.put("id", id);
        values.put("qsNumber", existing == null ? reference(offset, id) : existing);
        values.put("eTag", eTag(row));
        for (var person : PEOPLE.entrySet())
            readPerson(values, person.getKey(), row.get(person.getValue()));
        var sorted = new ArrayList<>(entries);
        sorted.sort(Comparator.comparing(ProgressLogEntry::getTs));
        values.put("progressLog", sorted);
        return json.convertValue(values, Issue.class);
    }

    private ProgressLogEntry toProgress(Map<String, Object> row) {
        var values = new LinkedHashMap<String, Object>();
        values.put("id", number(row.get("Id")));
        values.put("text", text(row.get("EntryText")));
        values.put("ts", text(row.get("Created")));
        readPerson(values, "author", row.get("Author"));
        return json.convertValue(values, ProgressLogEntry.class);
    }

    private void readPerson(Map<String, Object> values, String name, Object raw) {
        Map<?, ?> person = raw instanceof Map<?, ?> map ? map : Map.of();
        values.put(name, text(person.get("Title")));
        values.put(name + "Id", number(person.get("Id")));
        String email = text(person.get("EMail"));
        if (email.isBlank()) {
            String login = text(person.get("Name"));
            email = login.contains("|") ? login.substring(login.lastIndexOf('|') + 1) : "";
        }
        values.put(name + "Email", email);
    }

    private Map<String, Object> toFields(Map<String, Object> input) {
        var fields = new LinkedHashMap<String, Object>();
        for (var field : FIELDS.entrySet()) {
            String name = field.getKey();
            if (!input.containsKey(name)) continue;
            Object value = input.get(name);
            if (BOOLS.contains(name)) value = Boolean.TRUE.equals(value) ? "Yes" : "No";
            else if (CalendarDates.FIELDS.contains(name)) {
                String date = CalendarDates.normalize(text(value));
                value = date.isEmpty() ? null : date;
            } else if (TIMESTAMPS.contains(name) && text(value).isEmpty()) value = null;
            fields.put(field.getValue(), value);
        }
        if (input.containsKey("shortSummary")) fields.put("Title", text(input.get("shortSummary")));
        for (var person : PEOPLE.entrySet()) {
            String name = person.getKey();
            if (!input.containsKey(name)
                    && !input.containsKey(name + "Id")
                    && !input.containsKey(name + "Email")) continue;
            String email = text(input.get(name + "Email"));
            if (!email.isBlank()) {
                SharePointUser user;
                try {
                    user = client.ensureUser(email);
                } catch (org.springframework.web.client.HttpClientErrorException invalidUser) {
                    if (invalidUser.getStatusCode().value() != 400) throw invalidUser;
                    throw new IssueOperationException(
                            400,
                            "INVALID_PERSON",
                            "The Microsoft 365 identity could not be resolved for " + name + ".");
                }
                if (user.id() <= 0)
                    throw new IssueOperationException(
                            400,
                            "INVALID_PERSON",
                            "The Microsoft 365 identity could not be resolved.");
                fields.put(person.getValue() + "Id", user.id());
            } else if (input.containsKey(name + "Id")) {
                Long id = number(input.get(name + "Id"));
                fields.put(person.getValue() + "Id", id != null && id > 0 ? id : null);
            } else if (input.containsKey(name + "Email")
                    || (input.containsKey(name) && text(input.get(name)).isEmpty()))
                fields.put(person.getValue() + "Id", null);
            else
                throw new IssueOperationException(
                        400,
                        "INVALID_PERSON",
                        "A Microsoft 365 email or lookup ID is required for " + name + ".");
        }
        return fields;
    }

    public static String eTag(Map<String, Object> row) {
        Object value = row.get("odata.etag");
        if (value == null) value = row.get("@odata.etag");
        if (value == null && row.get("__metadata") instanceof Map<?, ?> metadata)
            value = metadata.get("etag");
        return value == null ? null : value.toString();
    }

    public static String text(Object value) {
        return value == null ? "" : value.toString();
    }

    public static Long number(Object value) {
        return value == null ? null : new BigDecimal(value.toString()).longValueExact();
    }
}
