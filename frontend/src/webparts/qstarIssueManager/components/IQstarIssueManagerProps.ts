import { IDataService } from "../services/IDataService";
import { ICheckResult } from "../services/ConnectionDiagnosticsService";

export interface IQstarIssueManagerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  dataService: IDataService;
  runConnectionDiagnostics: () => Promise<ICheckResult[]>;
}
