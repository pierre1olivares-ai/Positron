import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'QstarIssueManagerWebPartStrings';
import QstarIssueManager from './components/QstarIssueManager';
import { IQstarConnection, IQstarIssueManagerProps } from './components/IQstarIssueManagerProps';
import { SharePointDataService } from './services/SharePointDataService';
import { BackendApiClient } from './services/BackendApiClient';
import { BackendRoleResolver } from './services/BackendRoleResolver';
import { BackendApiDataService } from './services/BackendApiDataService';
import { ConnectionDiagnosticsService, ICheckResult } from './services/ConnectionDiagnosticsService';
import { BackendDiagnosticsService } from './services/BackendDiagnosticsService';
import { DEFAULT_ISSUES_LIST, DEFAULT_PROGRESS_LIST } from './services/fieldMap';
import { DevelopmentRoleResolver, SharePointRoleResolver } from './services/SharePointRoleResolver';
import { MockDataService } from './services/MockDataService';
import type { IDataService } from './services/IDataService';
import type { IRoleResolver } from './services/IRoleResolver';

export type DataSourceMode = 'backend' | 'sharepoint';

export interface IQstarIssueManagerWebPartProps {
  description: string;
  dataSourceMode: DataSourceMode;
  // Backend API mode (backend/service/) — the target architecture: SharePoint stays the real
  // store, reached server-side; this web part talks to the backend only.
  backendBaseUrl: string;
  backendResourceId: string;
  // SharePoint-direct mode (SharePointDataService.ts) — talks to SharePoint straight from the
  // browser via PnPjs. Works today with no backend deployment required; kept as the practical
  // default until the backend has somewhere real to run (see backend/service/README.md).
  siteUrl: string;
  issuesListName: string;
  progressListName: string;
  betaAccessMode: boolean;
}

export default class QstarIssueManagerWebPart extends BaseClientSideWebPart<IQstarIssueManagerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _connectionServices: {
    key: string;
    dataService: IDataService;
    roleResolver: IRoleResolver;
    diagnostics: { run(): Promise<ICheckResult[]> };
  } | undefined;

  public render(): void {
    const issuesListName = this.properties.issuesListName || DEFAULT_ISSUES_LIST;
    const progressListName = this.properties.progressListName || DEFAULT_PROGRESS_LIST;
    const siteUrl = this.properties.siteUrl ? this.properties.siteUrl.trim().replace(/\/+$/, '') : undefined;
    const mode: DataSourceMode = this.properties.dataSourceMode || 'sharepoint';
    const backendBaseUrl = (this.properties.backendBaseUrl || '').trim().replace(/\/+$/, '');
    const backendResourceId = (this.properties.backendResourceId || '').trim();
    const connection: IQstarConnection = {
      dataSourceMode: mode,
      backendBaseUrl: mode === 'backend' ? backendBaseUrl : undefined,
      siteUrl: siteUrl || this.context.pageContext.web.absoluteUrl,
      issuesListName,
      progressListName,
      betaAccessMode: !!this.properties.betaAccessMode
    };
    const connectionKey = JSON.stringify([
      mode,
      ...(mode === 'backend' ? [backendBaseUrl, backendResourceId] : [
        connection.siteUrl, issuesListName, progressListName, connection.betaAccessMode
      ]),
      this.context.isServedFromLocalhost,
      this.context.pageContext.user.email
    ]);

    if (!this._connectionServices || this._connectionServices.key !== connectionKey) {
      if (mode === 'backend') {
        const client = new BackendApiClient(this.context, backendResourceId, backendBaseUrl);
        this._connectionServices = {
          key: connectionKey,
          dataService: new BackendApiDataService(client),
          diagnostics: new BackendDiagnosticsService(client),
          roleResolver: new BackendRoleResolver(client, backendBaseUrl)
        };
      } else this._connectionServices = {
        key: connectionKey,
        dataService: this.context.isServedFromLocalhost
          ? new MockDataService()
          : new SharePointDataService(this.context, issuesListName, progressListName, siteUrl),
        diagnostics: new ConnectionDiagnosticsService(this.context, issuesListName, progressListName, siteUrl),
        roleResolver: this.context.isServedFromLocalhost
          ? new DevelopmentRoleResolver('admin')
          : new SharePointRoleResolver(this.context, undefined, siteUrl, connection.betaAccessMode)
      };
    }
    const { dataService, diagnostics, roleResolver } = this._connectionServices;

    const element: React.ReactElement<IQstarIssueManagerProps> = React.createElement(
      QstarIssueManager,
      {
        key: connectionKey,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        connection,
        dataService,
        roleResolver,
        runConnectionDiagnostics: (): Promise<ICheckResult[]> => diagnostics.run()
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneChoiceGroup('dataSourceMode', {
                  label: 'Data source',
                  options: [
                    { key: 'sharepoint', text: 'SharePoint direct (works today, no backend deployment needed)' },
                    { key: 'backend', text: 'Backend API (requires deployment and acceptance; disabled by default on the server)' }
                  ]
                })
              ]
            },
            {
              groupName: 'SharePoint direct settings',
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'SharePoint site URL (leave blank to use the site this web part is on)'
                }),
                PropertyPaneTextField('issuesListName', {
                  label: 'Issues list name',
                  value: DEFAULT_ISSUES_LIST
                }),
                PropertyPaneTextField('progressListName', {
                  label: 'Progress log list name',
                  value: DEFAULT_PROGRESS_LIST
                }),
                PropertyPaneToggle('betaAccessMode', {
                  label: 'Beta access mode (use existing site permissions)',
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            },
            {
              groupName: 'Backend API settings',
              groupFields: [
                PropertyPaneTextField('backendBaseUrl', {
                  label: 'Backend base URL, e.g. https://qstar.time-matters.com/api/v1'
                }),
                PropertyPaneTextField('backendResourceId', {
                  label: "Backend Entra application ID or resource URI"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
