import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { graph } from "@pnp/graph/presets/all";
import * as strings from "AdminDashboardWebPartStrings";
import AdminDashboard from "./components/AdminDashboard";
import { IAdminDashboardProps } from "./components/IAdminDashboardProps";

export interface IAdminDashboardWebPartProps {
  description: string;
}

export default class AdminDashboardWebPart extends BaseClientSideWebPart<IAdminDashboardWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAdminDashboardProps> =
      React.createElement(AdminDashboard, {
        description: this.properties.description,
        context: this.context,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
