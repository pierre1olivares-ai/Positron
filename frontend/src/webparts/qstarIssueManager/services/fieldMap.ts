// Internal SharePoint column names for the "Q-Star Issues" list and the
// "Q-Star Progress Log" child list. Must stay in sync with
// backend/sharepoint/qstar-sharepoint-graph-integration.md and the two
// provisioning scripts under backend/sharepoint/provisioning/.

export const ISSUE_FIELDS = {
  qsNumber: "QsNumber",
  shortSummary: "ShortSummary",
  description: "Description",
  immediateAction: "ImmediateAction",
  severity: "Severity",
  createdBy: "ReportedBy",
  reportDate: "ReportDate",
  departmentBU: "DepartmentBU",
  region: "Region",
  alreadyInContact: "AlreadyInContact",
  deviationType: "DeviationType",
  issueOrigin: "Origin",
  additionalComments: "AdditionalComments",

  followUp: "FollowUp",
  status: "Status",
  transformedInto: "TransformedInto",
  taskCreated: "TaskCreated",
  triaged: "Triaged",

  taskOwner: "TaskOwner",
  permissionedOwnerEmail: "PermissionedOwnerEmail",
  ownerBU: "EscalationBU",
  dueDate: "DueDate",

  rootCause: "RootCause",
  correctiveAction: "CorrectiveAction",
  implementationDate: "ImplementationDate",
  effectivenessCheck: "EffectivenessCheck",
  verifiedBy: "VerifiedBy",
  verifiedDate: "VerifiedDate",
  closedDate: "ClosedDate",
  closedAt: "ClosedAt",

  holdReason: "HoldReason",
  holdUntil: "HoldUntil",

  ownerUpdate: "OwnerUpdate",
  ownerUpdateAt: "OwnerUpdateAt",
  ownerUpdateText: "OwnerUpdateText",
} as const;

export const PROGRESS_FIELDS = {
  parentItemId: "ParentItemId",
  author: "Author",
  entryDate: "EntryDate",
  text: "EntryText",
} as const;

export const DEFAULT_ISSUES_LIST = "Q-Star Issues";
export const DEFAULT_PROGRESS_LIST = "Q-Star Progress Log";

/** Object.values isn't available under this project's ES5 lib target. */
export function objectValues<T extends Record<string, string>>(obj: T): string[] {
  return Object.keys(obj).map((k) => obj[k]);
}

export const ISSUE_FIELD_NAMES = objectValues(ISSUE_FIELDS);
export const PROGRESS_FIELD_NAMES = objectValues(PROGRESS_FIELDS);

export const ISSUE_PERSON_FIELD_NAMES: string[] = [
  ISSUE_FIELDS.createdBy,
  ISSUE_FIELDS.taskOwner,
  ISSUE_FIELDS.verifiedBy,
];

export const ISSUE_SCALAR_FIELD_NAMES = ISSUE_FIELD_NAMES.filter(
  (name) => ISSUE_PERSON_FIELD_NAMES.indexOf(name) < 0
);

export const ISSUE_SELECT_FIELD_NAMES = ISSUE_SCALAR_FIELD_NAMES.concat(
  ISSUE_PERSON_FIELD_NAMES.reduce<string[]>((all, name) => all.concat(
    `${name}/Id`,
    `${name}/Title`,
    `${name}/EMail`,
    `${name}/LoginName`
  ), [])
);

export const PROGRESS_SCALAR_FIELD_NAMES = PROGRESS_FIELD_NAMES.filter(
  (name) => name !== PROGRESS_FIELDS.author
);

export const PROGRESS_SELECT_FIELD_NAMES = PROGRESS_SCALAR_FIELD_NAMES.concat(
  `${PROGRESS_FIELDS.author}/Id`,
  `${PROGRESS_FIELDS.author}/Title`,
  `${PROGRESS_FIELDS.author}/EMail`,
  `${PROGRESS_FIELDS.author}/LoginName`
);
