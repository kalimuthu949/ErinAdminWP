import * as React from "react";
import styles from "./NonWellsFargoQuoteForm.module.scss";
import { INonWellsFargoQuoteFormProps } from "./INonWellsFargoQuoteFormProps";
import { escape } from "@microsoft/sp-lodash-subset";
import "../../../ExternalRef/css/style.css";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";

import NWFQuoteForm from "./NWFQuoteForm";

export default class NonWellsFargoQuoteForm extends React.Component<
  INonWellsFargoQuoteFormProps,
  {}
> {
  constructor(prop: INonWellsFargoQuoteFormProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<INonWellsFargoQuoteFormProps> {
    return <NWFQuoteForm spcontext={sp} />;
  }
}
