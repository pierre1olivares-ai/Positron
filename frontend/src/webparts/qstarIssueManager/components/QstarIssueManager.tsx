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
  public constructor(props: IQstarIssueManagerProps) {
    super(props);
    this.state = {};
  }

  public componentDidMount(): void {
    this.props.roleResolver.resolve()
      .then((roleResolution) => this.setState({ roleResolution, roleError: undefined }))
      .catch((error: Error) => this.setState({ roleError: error.message }));
  }

  public render(): React.ReactElement<IQstarIssueManagerProps> {
    const { roleResolution, roleError } = this.state;

    if (roleError) {
      return (
        <section className={styles.accessError} role="alert">
          <h2>Q-Star access could not be verified</h2>
          <p>Privileged features are disabled because your SharePoint role could not be resolved.</p>
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
          userDisplayName={this.props.userDisplayName}
          userEmail={this.props.userEmail}
          developmentMode={roleResolution.source === "development"}
          onRunDiagnostics={this.props.runConnectionDiagnostics}
        />
      </section>
    );
  }
}
