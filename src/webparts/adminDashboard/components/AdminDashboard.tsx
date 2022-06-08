import * as React from "react";
import styles from "./AdminDashboard.module.scss";
import { IAdminDashboardProps } from "./IAdminDashboardProps";
import { escape } from "@microsoft/sp-lodash-subset";
import App from "./App";
import "../../../ExternalRef/css/style.css";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";
export default class AdminDashboard extends React.Component<
  IAdminDashboardProps,
  {}
> {
  constructor(prop: IAdminDashboardProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<IAdminDashboardProps> {
    return <App spcontext={sp} graphcontext={graph} context={this.props.context}/>;
  }
}
