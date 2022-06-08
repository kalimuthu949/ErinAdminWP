import * as React from "react";
import styles from "./WellsFargoQuoteForm.module.scss";
import { IWellsFargoQuoteFormProps } from "./IWellsFargoQuoteFormProps";
import { escape } from "@microsoft/sp-lodash-subset";
import "../../../ExternalRef/css/style.css";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";
import WFQuoteForm from "./WFQuoteForm";
import Peoples from "./WFPeoples";

export default class WellsFargoQuoteForm extends React.Component<
  IWellsFargoQuoteFormProps,
  {}
> {
  constructor(prop: IWellsFargoQuoteFormProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<IWellsFargoQuoteFormProps> {
    return (
      <WFQuoteForm
        spcontext={sp}
        cont={this.props.context}
        graphcontext={graph}
      />
      // <Peoples spcontext={sp} cont={this.props.context} />
    );
  }
}
