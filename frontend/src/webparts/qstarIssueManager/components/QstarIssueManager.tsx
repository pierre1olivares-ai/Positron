import * as React from "react";

import styles from "./QstarIssueManager.module.scss";
import type { IQstarIssueManagerProps } from "./IQstarIssueManagerProps";
import type { IResolvedRole } from "../models/IRole";
import QstarPrototype from "./QstarPrototype";

export interface IQstarIssueManagerState {
  roleResolution?: IResolvedRole;
  roleError?: string;
}

export default class QstarIssueManager extends React.Component<IQstarIssueManagerProps, IQstarIssueManagerState> {
  private _mounted: boolean = false;

  public constructor(props: IQstarIssueManagerProps) {
    super(props);
    this.state = {};
  }

  public componentDidMount(): void {
    this._mounted = true;
    this.props.roleResolver.resolve()
      .then((roleResolution) => {
        if (roleResolution.source === "backend" && (!roleResolution.user || !roleResolution.connection)) {
          throw new Error("The backend did not return a verified user and connection. Access is disabled.");
        }
        if (this._mounted) this.setState({ roleResolution, roleError: undefined });
      })
      .catch((error: Error) => {
        if (this._mounted) this.setState({ roleError: error.message });
      });
  }

  public componentWillUnmount(): void {
    this._mounted = false;
  }

  public render(): React.ReactElement<IQstarIssueManagerProps> {
    const { roleResolution, roleError } = this.state;

    if (roleError) {
      return (
        <section className={styles.accessError} role="alert">
          <h2>Q-Star access could not be verified</h2>
          <p>Access is disabled because your Q-Star role could not be resolved.</p>
          <p className={styles.errorDetail}>{roleError}</p>
        </section>
      );
    }

    if (!roleResolution) {
      return <section className={styles.loading}>Checking your Q-Star access…</section>;
    }

    return (
      <section className={`${styles.qstarIssueManager} ${this.props.hasTeamsContext ? styles.teams : ""}`}>
        <QstarPrototype
          dataService={this.props.dataService}
          profile={roleResolution.role}
          userDisplayName={roleResolution.user?.displayName || this.props.userDisplayName}
          userEmail={roleResolution.user?.email || this.props.userEmail}
          connection={roleResolution.connection || this.props.connection}
          developmentMode={roleResolution.source === "development"}
          onRunDiagnostics={this.props.runConnectionDiagnostics}
        />
      </section>
    );
  }
}
