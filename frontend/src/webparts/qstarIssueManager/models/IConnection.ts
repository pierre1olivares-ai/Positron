export interface IQstarConnection {
  siteUrl: string;
  issuesListName: string;
  progressListName: string;
  betaAccessMode: boolean;
  dataSourceMode?: "sharepoint" | "backend";
  backendBaseUrl?: string;
}

export interface IBackendUser {
  userId: number;
  displayName: string;
  email: string;
}
