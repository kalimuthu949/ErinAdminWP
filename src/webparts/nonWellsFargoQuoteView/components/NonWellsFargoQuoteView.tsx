import * as React from "react";
import styles from "./NonWellsFargoQuoteView.module.scss";
import { INonWellsFargoQuoteViewProps } from "./INonWellsFargoQuoteViewProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";
import NWFQuoteView from "./NWFQuoteView";
import "../../../ExternalRef/css/style.css";

export default class NonWellsFargoQuoteView extends React.Component<
  INonWellsFargoQuoteViewProps,
  {}
> {
  constructor(prop: INonWellsFargoQuoteViewProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<INonWellsFargoQuoteViewProps> {
    return <NWFQuoteView spcontext={sp} context={this.props.context}/>;
  }
}
