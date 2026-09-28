package com.timematters.qstar.service;

import com.timematters.qstar.api.model.IssueATO;
import com.timematters.qstar.api.model.IssueCreateATO;
import com.timematters.qstar.api.model.IssuePatchATO;
import com.timematters.qstar.api.model.ProgressLogEntryATO;
import com.timematters.qstar.infrastructure.sharepoint.IssueFields;
import com.timematters.qstar.infrastructure.sharepoint.IssueRepository;
import com.timematters.qstar.model.Issue;
import com.timematters.qstar.model.ProgressLogEntry;
import com.timematters.qstar.model.mapper.IssueMapper;
import com.timematters.qstar.model.mapper.ProgressLogEntryMapper;
import java.time.LocalDate;
import java.time.OffsetDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

/**
 * Business logic for issues. Mirrors the intake/triage rules documented in
 * backend/sharepoint/qstar-sharepoint-graph-integration.md and implemented client-side in the
 * prototype (frontend/prototype/qstar-issue-manager.jsx) — this is a starting slice (intake
 * defaults + due-date SLA) rather than a full port of every rule (hold reminders, owner-update
 * flagging notifications, etc. still need to move here as the API is adopted).
 */
@Service
public class IssueService {

    private static final Map<String, Integer> SEVERITY_DUE_DAYS =
            Map.of(
                    "Critical", 7,
                    "High", 14,
                    "Medium", 30,
                    "Low", 60);

    private final IssueRepository issueRepository;
    private final IssueMapper issueMapper = new IssueMapper();
    private final ProgressLogEntryMapper progressLogEntryMapper = new ProgressLogEntryMapper();

    @Autowired
    public IssueService(IssueRepository issueRepository) {
        this.issueRepository = issueRepository;
    }

    public List<IssueATO> getAllIssues() {
        return issueRepository.findAll().stream().map(issueMapper::toATO).collect(Collectors.toList());
    }

    public IssueATO createIssue(IssueCreateATO create) {
        Issue issue = new Issue();
        issue.setShortSummary(create.getShortSummary());
        issue.setDescription(create.getDescription());
        issue.setImmediateAction(create.getImmediateAction());
        issue.setSeverity(create.getSeverity());
        issue.setCreatedBy(create.getCreatedBy());
        String reportDate = create.getReportDate() != null ? create.getReportDate() : LocalDate.now().toString();
        issue.setReportDate(reportDate);
        issue.setDepartmentBU(create.getDepartmentBU());
        issue.setRegion(create.getRegion());
        issue.setAlreadyInContact(create.getAlreadyInContact());
        issue.setDeviationType(create.getDeviationType());
        issue.setIssueOrigin(create.getIssueOrigin());
        issue.setAdditionalComments(create.getAdditionalComments());

        // Intake defaults — untriaged, no task yet, no due date until the QM triages it.
        issue.setTriaged(false);
        issue.setTaskCreated("No");
        issue.setQsNumber(nextQsNumber());

        Issue created = issueRepository.create(issue);
        return issueMapper.toATO(created);
    }

    public IssueATO updateIssue(Long id, IssuePatchATO patch) {
        Map<String, Object> fields = new HashMap<>();
        putIfNotNull(fields, IssueFields.TRIAGED, patch.getTriaged());
        putIfNotNull(fields, IssueFields.STATUS, patch.getStatus());
        putIfNotNull(fields, IssueFields.TASK_CREATED, patch.getTaskCreated());
        putIfNotNull(fields, IssueFields.TRANSFORMED_INTO, patch.getTransformedInto());
        putIfNotNull(fields, IssueFields.FOLLOW_UP, patch.getFollowUp());
        putIfNotNull(fields, IssueFields.TASK_OWNER, patch.getTaskOwner());
        putIfNotNull(fields, IssueFields.OWNER_BU, patch.getOwnerBU());
        putIfNotNull(fields, IssueFields.DUE_DATE, patch.getDueDate());
        putIfNotNull(fields, IssueFields.ROOT_CAUSE, patch.getRootCause());
        putIfNotNull(fields, IssueFields.CORRECTIVE_ACTION, patch.getCorrectiveAction());
        putIfNotNull(fields, IssueFields.IMPLEMENTATION_DATE, patch.getImplementationDate());
        putIfNotNull(fields, IssueFields.EFFECTIVENESS_CHECK, patch.getEffectivenessCheck());
        putIfNotNull(fields, IssueFields.VERIFIED_BY, patch.getVerifiedBy());
        putIfNotNull(fields, IssueFields.VERIFIED_DATE, patch.getVerifiedDate());
        putIfNotNull(fields, IssueFields.CLOSED_DATE, patch.getClosedDate());
        putIfNotNull(fields, IssueFields.CLOSED_AT, patch.getClosedAt());
        putIfNotNull(fields, IssueFields.HOLD_REASON, patch.getHoldReason());
        putIfNotNull(fields, IssueFields.HOLD_UNTIL, patch.getHoldUntil());
        putIfNotNull(fields, IssueFields.OWNER_UPDATE_AT, patch.getOwnerUpdateAt());
        putIfNotNull(fields, IssueFields.OWNER_UPDATE_TEXT, patch.getOwnerUpdateText());
        if (patch.getOwnerUpdate() != null) {
            fields.put(IssueFields.OWNER_UPDATE, patch.getOwnerUpdate() ? "Yes" : "No");
        }

        // Triage-time due-date default: severity SLA from the report date, unless the caller
        // (QM override) already supplied one via the patch.
        if (patch.getDueDate() == null && Boolean.TRUE.equals(patch.getTriaged())) {
            issueRepository.findAll().stream()
                    .filter(i -> id.equals(i.getId()))
                    .findFirst()
                    .ifPresent(
                            existing -> {
                                String defaultDueDate = defaultDueDateFor(existing);
                                if (defaultDueDate != null) {
                                    fields.put(IssueFields.DUE_DATE, defaultDueDate);
                                }
                            });
        }

        issueRepository.update(id, fields);
        return issueMapper.toATO(
                issueRepository.findAll().stream().filter(i -> id.equals(i.getId())).findFirst().orElseThrow());
    }

    public ProgressLogEntryATO addProgressLogEntry(Long issueId, ProgressLogEntryATO entryATO) {
        ProgressLogEntry entry = progressLogEntryMapper.fromATO(entryATO);
        if (entry.getTs() == null) {
            entry.setTs(OffsetDateTime.now().toString());
        }
        issueRepository.addProgressLogEntry(issueId, entry);
        return progressLogEntryMapper.toATO(entry);
    }

    private String defaultDueDateFor(Issue issue) {
        Integer days = SEVERITY_DUE_DAYS.get(issue.getSeverity());
        if (days == null || issue.getReportDate() == null) {
            return null;
        }
        return LocalDate.parse(issue.getReportDate()).plusDays(days).toString();
    }

    private Integer nextQsNumber() {
        return issueRepository.findAll().stream()
                        .map(Issue::getQsNumber)
                        .filter(java.util.Objects::nonNull)
                        .max(Integer::compareTo)
                        .orElse(1000)
                + 1;
    }

    private void putIfNotNull(Map<String, Object> map, String key, Object value) {
        if (value != null) {
            map.put(key, value);
        }
    }
}
