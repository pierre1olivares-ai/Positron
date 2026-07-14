import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

import {
  ISSUE_FIELDS,
  ISSUE_FIELD_NAMES,
  PROGRESS_FIELD_NAMES,
  DEFAULT_ISSUES_LIST,
  DEFAULT_PROGRESS_LIST,
} from "./fieldMap";
import {
  buildDiagnosticIssuePayload,
  buildDiagnosticProgressPayload,
} from "./diagnosticPayload";

export type CheckStatus = "pass" | "warn" | "fail";

export interface ICheckResult {
  name: string;
  status: CheckStatus;
  message: string;
}

const REQUIRED_STATUS_CHOICE = "Under Testing/Revision";
const RECOMMENDED_INDEXED_FIELDS = [ISSUE_FIELDS.status, ISSUE_FIELDS.triaged, ISSUE_FIELDS.dueDate];

/**
 * Runs a battery of live checks against the SharePoint site/lists this web
 * part is configured against. Meant to be triggered by a button in the app
 * (Admin/IT-settings tab) so the person deploying it in the real tenant can
 * self-verify the connection, since this can only be run under a real
 * signed-in session — it cannot be executed from outside the browser.
 */
export class ConnectionDiagnosticsService {
  private sp: SPFI;

  constructor(
    context: WebPartContext,
    private issuesListName: string = DEFAULT_ISSUES_LIST,
    private progressListName: string = DEFAULT_PROGRESS_LIST,
    siteUrl?: string
  ) {
    this.sp = siteUrl ? spfi(siteUrl).using(SPFx(context)) : spfi().using(SPFx(context));
  }

  public async run(): Promise<ICheckResult[]> {
    const results: ICheckResult[] = [];

    results.push(await this.checkSiteAccess());
    results.push(await this.checkCurrentUser());
    const issuesListOk = await this.checkListExists(this.issuesListName, results);
    if (issuesListOk) {
      await this.checkRequiredFields(this.issuesListName, ISSUE_FIELD_NAMES, results);
      await this.checkStatusChoiceValue(results);
      await this.checkIndexedFields(results);
    }
    const progressListOk = await this.checkListExists(this.progressListName, results);
    if (progressListOk) {
      await this.checkRequiredFields(this.progressListName, PROGRESS_FIELD_NAMES, results);
    }
    await this.checkConfigList(results);

    if (issuesListOk && progressListOk) {
      await this.checkRoundTrip(results);
    } else {
      results.push({
        name: "Round-trip write test",
        status: "fail",
        message: "Skipped — required lists are missing. Run the provisioning script first.",
      });
    }

    return results;
  }

  private async checkSiteAccess(): Promise<ICheckResult> {
    try {
      const web = await this.sp.web.select("Title", "Url")();
      return { name: "Site access", status: "pass", message: `Connected to "${web.Title}" (${web.Url}).` };
    } catch (e) {
      return { name: "Site access", status: "fail", message: `Could not reach the site: ${errorMessage(e)}` };
    }
  }

  private async checkCurrentUser(): Promise<ICheckResult> {
    try {
      const user = await this.sp.web.currentUser();
      return {
        name: "Signed-in user",
        status: "pass",
        message: `Resolved as ${user.Title} (${user.Email || user.LoginName}).`,
      };
    } catch (e) {
      return { name: "Signed-in user", status: "fail", message: `Could not resolve current user: ${errorMessage(e)}` };
    }
  }

  private async checkListExists(listName: string, results: ICheckResult[]): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).select("Id")();
      results.push({ name: `List "${listName}"`, status: "pass", message: "Found." });
      return true;
    } catch (e) {
      results.push({
        name: `List "${listName}"`,
        status: "fail",
        message: `Not found — run the provisioning script. (${errorMessage(e)})`,
      });
      return false;
    }
  }

  private async checkRequiredFields(listName: string, internalNames: string[], results: ICheckResult[]): Promise<void> {
    try {
      const fields = await this.sp.web.lists
        .getByTitle(listName)
        .fields.select("InternalName")();
      const present = new Set(fields.map((f) => f.InternalName));
      const missing = internalNames.filter((n) => !present.has(n));
      if (missing.length === 0) {
        results.push({ name: `Columns on "${listName}"`, status: "pass", message: `All ${internalNames.length} expected columns present.` });
      } else {
        results.push({
          name: `Columns on "${listName}"`,
          status: "fail",
          message: `Missing columns: ${missing.join(", ")}. Re-run the provisioning script.`,
        });
      }
    } catch (e) {
      results.push({ name: `Columns on "${listName}"`, status: "fail", message: errorMessage(e) });
    }
  }

  private async checkStatusChoiceValue(results: ICheckResult[]): Promise<void> {
    try {
      const field = await this.sp.web.lists
        .getByTitle(this.issuesListName)
        .fields.getByInternalNameOrTitle(ISSUE_FIELDS.status)
        .select("Choices")();
      const choices: string[] = (field as unknown as { Choices: string[] }).Choices || [];
      if (choices.indexOf(REQUIRED_STATUS_CHOICE) >= 0) {
        results.push({ name: "Status choice values", status: "pass", message: "\"Under Testing/Revision\" present." });
      } else {
        results.push({
          name: "Status choice values",
          status: "fail",
          message: `"${REQUIRED_STATUS_CHOICE}" is missing from the Status choice list — NC effectiveness-test tracking will break.`,
        });
      }
    } catch (e) {
      results.push({ name: "Status choice values", status: "warn", message: `Could not verify: ${errorMessage(e)}` });
    }
  }

  private async checkIndexedFields(results: ICheckResult[]): Promise<void> {
    try {
      const fields = await this.sp.web.lists
        .getByTitle(this.issuesListName)
        .fields.select("InternalName", "Indexed")();
      const byName = new Map(fields.map((f) => [f.InternalName, f.Indexed]));
      const notIndexed = RECOMMENDED_INDEXED_FIELDS.filter((n) => !byName.get(n));
      if (notIndexed.length === 0) {
        results.push({ name: "Recommended indexes", status: "pass", message: "Status, Triaged and DueDate are all indexed." });
      } else {
        results.push({
          name: "Recommended indexes",
          status: "warn",
          message: `Not indexed: ${notIndexed.join(", ")}. Index these for the reminder flow's filtered queries to stay reliable at scale.`,
        });
      }
    } catch (e) {
      results.push({ name: "Recommended indexes", status: "warn", message: `Could not verify: ${errorMessage(e)}` });
    }
  }

  private async checkConfigList(results: ICheckResult[]): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle("Q-Star Config").select("Id")();
      results.push({ name: 'List "Q-Star Config"', status: "pass", message: "Found." });
    } catch {
      results.push({
        name: 'List "Q-Star Config"',
        status: "warn",
        message: "Not found — optional, only needed to persist IT-settings values centrally instead of per-browser.",
      });
    }
  }

  /** Creates, updates, and deletes a throwaway item in both lists to prove write/delete permissions end-to-end. */
  private async checkRoundTrip(results: ICheckResult[]): Promise<void> {
    let issueId: number | undefined;
    let progressId: number | undefined;
    try {
      const currentUser = await this.sp.web.currentUser.select("Id")();
      const issues = this.sp.web.lists.getByTitle(this.issuesListName).items;
      const created = await issues.add(buildDiagnosticIssuePayload());
      issueId = created.Id as number;
      const currentIssueId: number = issueId;

      await issues.getById(currentIssueId).update({ [ISSUE_FIELDS.followUp]: "diagnostic update" });

      const progress = this.sp.web.lists.getByTitle(this.progressListName).items;
      const createdEntry = await progress.add(buildDiagnosticProgressPayload(currentIssueId, currentUser.Id));
      progressId = createdEntry.Id as number;
      await progress.getById(progressId).delete();
      progressId = undefined;
      await issues.getById(currentIssueId).delete();
      issueId = undefined;

      results.push({
        name: "Round-trip write test",
        status: "pass",
        message: "Created, updated, and deleted a test item in both lists successfully.",
      });
    } catch (e) {
      results.push({ name: "Round-trip write test", status: "fail", message: errorMessage(e) });
      if (progressId !== undefined) {
        try {
          await this.sp.web.lists.getByTitle(this.progressListName).items.getById(progressId).delete();
        } catch {
          results.push({
            name: "Round-trip progress cleanup",
            status: "warn",
            message: `Left behind a progress test item (Id ${progressId}) in "${this.progressListName}" — delete it manually.`,
          });
        }
      }
      if (issueId !== undefined) {
        try {
          await this.sp.web.lists.getByTitle(this.issuesListName).items.getById(issueId).delete();
        } catch {
          results.push({
            name: "Round-trip cleanup",
            status: "warn",
            message: `Left behind a test item (Id ${issueId}) in "${this.issuesListName}" — delete it manually.`,
          });
        }
      }
    }
  }
}

function errorMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  return String(e);
}
