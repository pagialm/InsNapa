import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "InsuranceNapaWebPartStrings";
import InsuranceNapa from "./components/InsuranceNapa";
import { IInsuranceNapaProps } from "./components/IInsuranceNapaProps";

export interface IInsuranceNapaWebPartProps {
  description: string;
}

export default class InsuranceNapaWebPart extends BaseClientSideWebPart<IInsuranceNapaWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IInsuranceNapaProps> = React.createElement(
      InsuranceNapa,
      {
        description: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getdataVersion(): Version {
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
