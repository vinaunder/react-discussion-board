import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "ReactDiscussionBoardWebPartStrings";
import {ReactDiscussionBoard} from "./components/ReactDiscussionBoard";
import { IReactDiscussionBoardProps } from "./components/IReactDiscussionBoardProps";

import { sp } from "@pnp/sp";

export interface IReactDiscussionBoardWebPartProps {
  description: string;
  siteurl: string;
  listname: string;
}

export default class ReactDiscussionBoardWebPart extends BaseClientSideWebPart<IReactDiscussionBoardWebPartProps> {
  protected async onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context,
    });
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IReactDiscussionBoardProps> = React.createElement(
      ReactDiscussionBoard,
      {
        siteurl: this.properties.siteurl,
        listname: this.properties.listname,
        description: this.properties.description,
      }
    );

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
                PropertyPaneTextField("siteurl", {
                  label: "Site Url",
                }),
                PropertyPaneTextField("listname", {
                  label: "List Name",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
