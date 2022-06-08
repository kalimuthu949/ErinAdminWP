import * as React from "react";
import styles from "./WellsFargoQuoteView.module.scss";
import { IWellsFargoQuoteViewProps } from "./IWellsFargoQuoteViewProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";
import WFQuoteView from "./WFQuoteView";
import "../../../ExternalRef/css/style.css";
export default class WellsFargoQuoteView extends React.Component<
  IWellsFargoQuoteViewProps,
  {}
> {
  constructor(prop: IWellsFargoQuoteViewProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<IWellsFargoQuoteViewProps> {
    return <WFQuoteView spcontext={sp} graphcontext={graph} context={this.props.context}/>;
  }
}
