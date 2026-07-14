import { IDataService } from "../services/IDataService";
import { ICheckResult } from "../services/ConnectionDiagnosticsService";
import { IRoleResolver } from "../services/IRoleResolver";

export interface IQstarIssueManagerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  dataService: IDataService;
  roleResolver: IRoleResolver;
  runConnectionDiagnostics: () => Promise<ICheckResult[]>;
}
