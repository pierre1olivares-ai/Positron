import { IDataService } from "../services/IDataService";
import { ICheckResult } from "../services/ConnectionDiagnosticsService";
import { IRoleResolver } from "../services/IRoleResolver";

export interface IQstarConnection {
  siteUrl: string;
  issuesListName: string;
  progressListName: string;
  betaAccessMode: boolean;
}

export interface IQstarIssueManagerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  connection: IQstarConnection;
  dataService: IDataService;
  roleResolver: IRoleResolver;
  runConnectionDiagnostics: () => Promise<ICheckResult[]>;
}
