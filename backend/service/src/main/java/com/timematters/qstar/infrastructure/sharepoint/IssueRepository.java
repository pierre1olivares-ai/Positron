package com.timematters.qstar.infrastructure.sharepoint;

import com.timematters.qstar.model.Issue;
import com.timematters.qstar.model.ProgressLogEntry;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Repository;

/**
 * Reads/writes Q-Star issues against the "Q-Star Issues" and "Q-Star Progress Log" SharePoint
 * lists via Microsoft Graph (app-only, Sites.Selected). See
 * backend/sharepoint/qstar-sharepoint-graph-integration.md for the full field/column design this
 * mirrors.
 */
@Repository
public class IssueRepository {

    @Value("${qstar.sharepoint.issues-list}")
    private String issuesListName;

    @Value("${qstar.sharepoint.progress-list}")
    private String progressListName;

    private final SharePointGraphClient graphClient;

    @Autowired
    public IssueRepository(SharePointGraphClient graphClient) {
        this.graphClient = graphClient;
    }

    public List<Issue> findAll() {
        String issuesListId = graphClient.getListId(issuesListName);
        String progressListId = graphClient.getListId(progressListName);

        List<Map<String, Object>> rawIssues = graphClient.getAllItems(issuesListId);
        List<Map<String, Object>> rawProgress = graphClient.getAllItems(progressListId);

        Map<String, List<ProgressLogEntry>> logsByParentId = new HashMap<>();
        for (Map<String, Object> p : rawProgress) {
            String parentId = stringOf(p.get(IssueFields.Progress.PARENT_ITEM_ID));
            logsByParentId
                    .computeIfAbsent(parentId, k -> new ArrayList<>())
                    .add(
                            new ProgressLogEntry(
                                    stringOf(p.get(IssueFields.Progress.ENTRY_DATE)),
                                    stringOf(p.get(IssueFields.Progress.AUTHOR)),
                                    stringOf(p.get(IssueFields.Progress.TEXT))));
        }

        return rawIssues.stream()
                .map(fields -> toIssue(fields, logsByParentId.getOrDefault(stringOf(fields.get("Id")), List.of())))
                .collect(Collectors.toList());
    }

    public Issue create(Issue issue) {
        String issuesListId = graphClient.getListId(issuesListName);
        Map<String, Object> created = graphClient.createItem(issuesListId, toFields(issue));
        return toIssue(created, List.of());
    }

    public void update(Long id, Map<String, Object> patchFields) {
        String issuesListId = graphClient.getListId(issuesListName);
        graphClient.updateItemFields(issuesListId, id.toString(), patchFields);
    }

    public void addProgressLogEntry(Long issueId, ProgressLogEntry entry) {
        String progressListId = graphClient.getListId(progressListName);
        Map<String, Object> fields = new HashMap<>();
        fields.put(IssueFields.Progress.PARENT_ITEM_ID, issueId);
        fields.put(IssueFields.Progress.AUTHOR, entry.getAuthor());
        fields.put(IssueFields.Progress.ENTRY_DATE, entry.getTs());
        fields.put(IssueFields.Progress.TEXT, entry.getText());
        graphClient.createItem(progressListId, fields);
    }

    private Issue toIssue(Map<String, Object> f, List<ProgressLogEntry> progressLog) {
        Issue issue = new Issue();
        issue.setId(longOf(f.get("Id")));
        issue.setQsNumber(intOf(f.get(IssueFields.QS_NUMBER)));
        issue.setTriaged("Yes".equals(f.get(IssueFields.TRIAGED)));
        issue.setStatus(stringOf(f.get(IssueFields.STATUS)));
        issue.setTaskCreated((String) f.getOrDefault(IssueFields.TASK_CREATED, "No"));
        issue.setTransformedInto(stringOf(f.get(IssueFields.TRANSFORMED_INTO)));

        issue.setShortSummary(stringOf(f.get(IssueFields.SHORT_SUMMARY)));
        issue.setDescription(stringOf(f.get(IssueFields.DESCRIPTION)));
        issue.setImmediateAction(stringOf(f.get(IssueFields.IMMEDIATE_ACTION)));
        issue.setSeverity(stringOf(f.get(IssueFields.SEVERITY)));
        issue.setCreatedBy(stringOf(f.get(IssueFields.CREATED_BY)));
        issue.setReportDate(stringOf(f.get(IssueFields.REPORT_DATE)));
        issue.setDepartmentBU(stringOf(f.get(IssueFields.DEPARTMENT_BU)));
        issue.setRegion(stringOf(f.get(IssueFields.REGION)));
        issue.setAlreadyInContact((String) f.getOrDefault(IssueFields.ALREADY_IN_CONTACT, "No"));
        issue.setDeviationType(stringOf(f.get(IssueFields.DEVIATION_TYPE)));
        issue.setIssueOrigin(stringOf(f.get(IssueFields.ISSUE_ORIGIN)));
        issue.setAdditionalComments(stringOf(f.get(IssueFields.ADDITIONAL_COMMENTS)));

        issue.setFollowUp(stringOf(f.get(IssueFields.FOLLOW_UP)));
        issue.setTaskOwner(stringOf(f.get(IssueFields.TASK_OWNER)));
        String ownerBU = stringOf(f.get(IssueFields.OWNER_BU));
        issue.setOwnerBU(ownerBU != null ? ownerBU : stringOf(f.get(IssueFields.DEPARTMENT_BU)));
        issue.setDueDate(stringOf(f.get(IssueFields.DUE_DATE)));

        issue.setRootCause(stringOf(f.get(IssueFields.ROOT_CAUSE)));
        issue.setCorrectiveAction(stringOf(f.get(IssueFields.CORRECTIVE_ACTION)));
        issue.setImplementationDate(stringOf(f.get(IssueFields.IMPLEMENTATION_DATE)));
        issue.setEffectivenessCheck(stringOf(f.get(IssueFields.EFFECTIVENESS_CHECK)));
        issue.setVerifiedBy(stringOf(f.get(IssueFields.VERIFIED_BY)));
        issue.setVerifiedDate(stringOf(f.get(IssueFields.VERIFIED_DATE)));
        issue.setClosedDate(stringOf(f.get(IssueFields.CLOSED_DATE)));
        issue.setClosedAt(stringOf(f.get(IssueFields.CLOSED_AT)));

        issue.setHoldReason(stringOf(f.get(IssueFields.HOLD_REASON)));
        issue.setHoldUntil(stringOf(f.get(IssueFields.HOLD_UNTIL)));

        issue.setOwnerUpdate("Yes".equals(f.get(IssueFields.OWNER_UPDATE)));
        issue.setOwnerUpdateAt(stringOf(f.get(IssueFields.OWNER_UPDATE_AT)));
        issue.setOwnerUpdateText(stringOf(f.get(IssueFields.OWNER_UPDATE_TEXT)));

        issue.setProgressLog(progressLog);
        return issue;
    }

    /** Converts every non-null field on {@code issue} to SharePoint internal field names. */
    private Map<String, Object> toFields(Issue issue) {
        Map<String, Object> fields = new HashMap<>();
        putIfNotNull(fields, IssueFields.QS_NUMBER, issue.getQsNumber());
        putIfNotNull(fields, IssueFields.TRIAGED, boolToYesNo(issue.getTriaged()));
        putIfNotNull(fields, IssueFields.STATUS, issue.getStatus());
        putIfNotNull(fields, IssueFields.TASK_CREATED, issue.getTaskCreated());
        putIfNotNull(fields, IssueFields.TRANSFORMED_INTO, issue.getTransformedInto());

        putIfNotNull(fields, IssueFields.SHORT_SUMMARY, issue.getShortSummary());
        putIfNotNull(fields, IssueFields.DESCRIPTION, issue.getDescription());
        putIfNotNull(fields, IssueFields.IMMEDIATE_ACTION, issue.getImmediateAction());
        putIfNotNull(fields, IssueFields.SEVERITY, issue.getSeverity());
        putIfNotNull(fields, IssueFields.CREATED_BY, issue.getCreatedBy());
        putIfNotNull(fields, IssueFields.REPORT_DATE, issue.getReportDate());
        putIfNotNull(fields, IssueFields.DEPARTMENT_BU, issue.getDepartmentBU());
        putIfNotNull(fields, IssueFields.REGION, issue.getRegion());
        putIfNotNull(fields, IssueFields.ALREADY_IN_CONTACT, issue.getAlreadyInContact());
        putIfNotNull(fields, IssueFields.DEVIATION_TYPE, issue.getDeviationType());
        putIfNotNull(fields, IssueFields.ISSUE_ORIGIN, issue.getIssueOrigin());
        putIfNotNull(fields, IssueFields.ADDITIONAL_COMMENTS, issue.getAdditionalComments());

        putIfNotNull(fields, IssueFields.FOLLOW_UP, issue.getFollowUp());
        putIfNotNull(fields, IssueFields.TASK_OWNER, issue.getTaskOwner());
        putIfNotNull(fields, IssueFields.OWNER_BU, issue.getOwnerBU());
        putIfNotNull(fields, IssueFields.DUE_DATE, issue.getDueDate());

        putIfNotNull(fields, IssueFields.ROOT_CAUSE, issue.getRootCause());
        putIfNotNull(fields, IssueFields.CORRECTIVE_ACTION, issue.getCorrectiveAction());
        putIfNotNull(fields, IssueFields.IMPLEMENTATION_DATE, issue.getImplementationDate());
        putIfNotNull(fields, IssueFields.EFFECTIVENESS_CHECK, issue.getEffectivenessCheck());
        putIfNotNull(fields, IssueFields.VERIFIED_BY, issue.getVerifiedBy());
        putIfNotNull(fields, IssueFields.VERIFIED_DATE, issue.getVerifiedDate());
        putIfNotNull(fields, IssueFields.CLOSED_DATE, issue.getClosedDate());
        putIfNotNull(fields, IssueFields.CLOSED_AT, issue.getClosedAt());

        putIfNotNull(fields, IssueFields.HOLD_REASON, issue.getHoldReason());
        putIfNotNull(fields, IssueFields.HOLD_UNTIL, issue.getHoldUntil());

        if (issue.getOwnerUpdate() != null) {
            fields.put(IssueFields.OWNER_UPDATE, boolToYesNo(issue.getOwnerUpdate()));
        }
        putIfNotNull(fields, IssueFields.OWNER_UPDATE_AT, issue.getOwnerUpdateAt());
        putIfNotNull(fields, IssueFields.OWNER_UPDATE_TEXT, issue.getOwnerUpdateText());
        return fields;
    }

    private void putIfNotNull(Map<String, Object> map, String key, Object value) {
        if (value != null) {
            map.put(key, value);
        }
    }

    private String boolToYesNo(Boolean value) {
        return Boolean.TRUE.equals(value) ? "Yes" : "No";
    }

    private String stringOf(Object value) {
        return value == null ? null : value.toString();
    }

    private Integer intOf(Object value) {
        return value == null ? null : Integer.valueOf(value.toString());
    }

    private Long longOf(Object value) {
        return value == null ? null : Long.valueOf(value.toString());
    }
}
