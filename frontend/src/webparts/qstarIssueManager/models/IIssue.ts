// Keep SharePoint internal field names and serialization in services/fieldMap.ts
// and services/SharePointDataService.ts, outside the UI's issue model.

export type Severity = "Critical" | "High" | "Medium" | "Low";

export type IssueStatus =
  | "Created"
  | "In Progress"
  | "Under Testing/Revision"
  | "On Hold"
  | "Closed"
  | "Rejected";

export type TransformedInto =
  | "OFI"
  | "NC Minor"
  | "NC Major"
  | "Only sent to Dept/BU for Action";

export type YesNo = "Yes" | "No";

export interface IProgressLogEntry {
  id?: number;
  ts: string;
  author: string;
  authorId?: number;
  authorEmail?: string;
  text: string;
  saveWarning?: string;
}

export interface IIssue {
  qsNumber: number;
  id: number;
  /** SharePoint version used for an optimistic-concurrency update. */
  eTag?: string;
  /** The create was accepted; do not repeat it when reference synchronization or readback fails. */
  saveWarning?: string;
  triaged: boolean;
  status: IssueStatus | undefined;
  taskCreated: YesNo;
  transformedInto: TransformedInto | undefined;

  shortSummary: string;
  description: string;
  immediateAction: string;
  severity: Severity;
  createdBy: string;
  createdById?: number;
  createdByEmail?: string;
  reportDate: string;
  departmentBU: string;
  region: string;
  alreadyInContact: YesNo;
  deviationType: string;
  issueOrigin: string;
  additionalComments: string;

  followUp: string;
  taskOwner: string;
  taskOwnerId?: number;
  taskOwnerEmail?: string;
  ownerBU: string;
  dueDate: string;

  rootCause: string;
  correctiveAction: string;
  implementationDate: string;
  effectivenessCheck: string;
  verifiedBy: string;
  verifiedById?: number;
  verifiedByEmail?: string;
  verifiedDate: string;
  closedDate: string;
  closedAt: string;

  holdReason: string;
  holdUntil: string;

  ownerUpdate: boolean;
  ownerUpdateAt: string;
  ownerUpdateText: string;
  reminderCycle: string;

  attachments: unknown[];
  progressLog: IProgressLogEntry[];
}
