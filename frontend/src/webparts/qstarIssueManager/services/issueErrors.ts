import { IIssue } from "../models/IIssue";
import type { IAcceptedReceipt } from "../domain/acceptedWriteRecovery";
import { ACCEPTED_RECEIPT_MESSAGE } from "../domain/acceptedWriteRecovery";

export class AcceptedWriteError extends Error {
  public readonly saved = true;
  public readonly code = "ACCEPTED_WRITE";

  constructor(
    public readonly operation: "create" | "progress",
    public readonly identity: Pick<IAcceptedReceipt, "issueId" | "qsNumber" | "entryId"> = {}
  ) {
    super(ACCEPTED_RECEIPT_MESSAGE);
    this.name = "AcceptedWriteError";
    (Object as ObjectConstructor & { setPrototypeOf(target: object, prototype: object): object })
      .setPrototypeOf(this, AcceptedWriteError.prototype);
  }
}

export class IssueConflictError extends Error {
  public readonly code = "ISSUE_CONFLICT";

  constructor(public readonly issueId: number, public readonly freshIssue?: IIssue) {
    super("This issue changed since you opened it. Review the latest version before saving again.");
    this.name = "IssueConflictError";
    (Object as ObjectConstructor & { setPrototypeOf(target: object, prototype: object): object })
      .setPrototypeOf(this, IssueConflictError.prototype);
  }
}

export class IssueRefreshError extends Error {
  public readonly code = "ISSUE_REFRESH";
  public readonly saved = true;

  constructor(public readonly issueId: number) {
    super("Your changes were saved, but the latest issue could not be reloaded. Refresh before editing again; do not submit the change again.");
    this.name = "IssueRefreshError";
    (Object as ObjectConstructor & { setPrototypeOf(target: object, prototype: object): object })
      .setPrototypeOf(this, IssueRefreshError.prototype);
  }
}
