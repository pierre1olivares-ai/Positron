import * as React from 'react';
import styles from './QstarIssueManager.module.scss';
import type { IQstarIssueManagerProps } from './IQstarIssueManagerProps';
import type { ICheckResult } from '../services/ConnectionDiagnosticsService';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IQstarIssueManagerState {
  diagnosticsRunning: boolean;
  diagnosticsResults: ICheckResult[] | undefined;
  diagnosticsError: string | undefined;
}

const STATUS_ICON: Record<ICheckResult['status'], string> = {
  pass: '✅',
  warn: '⚠️',
  fail: '❌',
};

function statusClassName(status: ICheckResult['status']): string {
  switch (status) {
    case 'pass': return styles.diagnostic_pass;
    case 'warn': return styles.diagnostic_warn;
    case 'fail': return styles.diagnostic_fail;
    default: return '';
  }
}

export default class QstarIssueManager extends React.Component<IQstarIssueManagerProps, IQstarIssueManagerState> {
  constructor(props: IQstarIssueManagerProps) {
    super(props);
    this.state = { diagnosticsRunning: false, diagnosticsResults: undefined, diagnosticsError: undefined };
  }

  private runDiagnostics = (): void => {
    this.setState({ diagnosticsRunning: true, diagnosticsError: undefined });
    this.props
      .runConnectionDiagnostics()
      .then((results) => this.setState({ diagnosticsRunning: false, diagnosticsResults: results }))
      .catch((e: Error) => this.setState({ diagnosticsRunning: false, diagnosticsError: e.message }));
  };

  public render(): React.ReactElement<IQstarIssueManagerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    const { diagnosticsRunning, diagnosticsResults, diagnosticsError } = this.state;

    return (
      <section className={`${styles.qstarIssueManager} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>

        <div className={styles.diagnostics}>
          <h3>Connection diagnostics</h3>
          <p>Run this once the &quot;Q-Star Issues&quot; / &quot;Q-Star Progress Log&quot; lists are provisioned to confirm this web part can read and write them under your own permissions.</p>
          <button onClick={this.runDiagnostics} disabled={diagnosticsRunning}>
            {diagnosticsRunning ? 'Running…' : 'Run Connection Test'}
          </button>
          {diagnosticsError && <p className={styles.diagnosticFail}>{diagnosticsError}</p>}
          {diagnosticsResults && (
            <ul className={styles.diagnosticList}>
              {diagnosticsResults.map((r, idx) => (
                <li key={idx} className={statusClassName(r.status)}>
                  <span>{STATUS_ICON[r.status]}</span> <strong>{r.name}</strong> — {r.message}
                </li>
              ))}
            </ul>
          )}
        </div>
      </section>
    );
  }
}
