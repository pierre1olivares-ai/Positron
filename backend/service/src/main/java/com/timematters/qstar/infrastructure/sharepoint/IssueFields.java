package com.timematters.qstar.infrastructure.sharepoint;

/**
 * Internal SharePoint column names for the "Q-Star Issues" list. Mirrors
 * frontend/src/webparts/qstarIssueManager/services/fieldMap.ts (ISSUE_FIELDS) and
 * backend/sharepoint/qstar-sharepoint-graph-integration.md — keep all three in sync. Person columns
 * are stored as text + a companion "*Email" column (the -PersonAsText provisioning option),
 * matching the prototype's plain-string model.
 */
public final class IssueFields {
    private IssueFields() {}

    public static final String QS_NUMBER = "QsNumber";
    public static final String SHORT_SUMMARY = "ShortSummary";
    public static final String DESCRIPTION = "Description";
    public static final String IMMEDIATE_ACTION = "ImmediateAction";
    public static final String SEVERITY = "Severity";
    public static final String CREATED_BY = "ReportedBy";
    public static final String REPORT_DATE = "ReportDate";
    public static final String DEPARTMENT_BU = "DepartmentBU";
    public static final String REGION = "Region";
    public static final String ALREADY_IN_CONTACT = "AlreadyInContact";
    public static final String DEVIATION_TYPE = "DeviationType";
    public static final String ISSUE_ORIGIN = "Origin";
    public static final String ADDITIONAL_COMMENTS = "AdditionalComments";

    public static final String FOLLOW_UP = "FollowUp";
    public static final String STATUS = "Status";
    public static final String TRANSFORMED_INTO = "TransformedInto";
    public static final String TASK_CREATED = "TaskCreated";
    public static final String TRIAGED = "Triaged";

    public static final String TASK_OWNER = "TaskOwner";
    public static final String OWNER_BU = "EscalationBU";
    public static final String DUE_DATE = "DueDate";

    public static final String ROOT_CAUSE = "RootCause";
    public static final String CORRECTIVE_ACTION = "CorrectiveAction";
    public static final String IMPLEMENTATION_DATE = "ImplementationDate";
    public static final String EFFECTIVENESS_CHECK = "EffectivenessCheck";
    public static final String VERIFIED_BY = "VerifiedBy";
    public static final String VERIFIED_DATE = "VerifiedDate";
    public static final String CLOSED_DATE = "ClosedDate";
    public static final String CLOSED_AT = "ClosedAt";

    public static final String HOLD_REASON = "HoldReason";
    public static final String HOLD_UNTIL = "HoldUntil";

    public static final String OWNER_UPDATE = "OwnerUpdate";
    public static final String OWNER_UPDATE_AT = "OwnerUpdateAt";
    public static final String OWNER_UPDATE_TEXT = "OwnerUpdateText";

    public static final class Progress {
        private Progress() {}

        public static final String PARENT_ITEM_ID = "ParentItemId";
        public static final String AUTHOR = "Author";
        public static final String ENTRY_DATE = "EntryDate";
        public static final String TEXT = "EntryText";
    }
}
