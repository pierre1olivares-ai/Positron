package com.timematters.qstar.model;

import java.util.ArrayList;
import java.util.List;
import lombok.Data;

/**
 * Internal representation of a Q-Star issue. Field names deliberately mirror
 * frontend/src/webparts/qstarIssueManager/models/IIssue.ts one-for-one, which itself mirrors the
 * prototype's in-memory issue object — so the ATO mapping (generated from
 * api-contract/contract.yaml) and the SharePoint field mapping (IssueFields) both stay
 * straightforward 1:1 lookups.
 */
@Data
public class Issue {
    private Long id;
    private Integer qsNumber;
    private Boolean triaged;
    private String status;
    private String taskCreated;
    private String transformedInto;

    private String shortSummary;
    private String description;
    private String immediateAction;
    private String severity;
    private String createdBy;
    private String reportDate;
    private String departmentBU;
    private String region;
    private String alreadyInContact;
    private String deviationType;
    private String issueOrigin;
    private String additionalComments;

    private String followUp;
    private String taskOwner;
    private String ownerBU;
    private String dueDate;

    private String rootCause;
    private String correctiveAction;
    private String implementationDate;
    private String effectivenessCheck;
    private String verifiedBy;
    private String verifiedDate;
    private String closedDate;
    private String closedAt;

    private String holdReason;
    private String holdUntil;

    private Boolean ownerUpdate;
    private String ownerUpdateAt;
    private String ownerUpdateText;

    private List<ProgressLogEntry> progressLog = new ArrayList<>();
}
