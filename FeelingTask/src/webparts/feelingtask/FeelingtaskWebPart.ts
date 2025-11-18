import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// Add PnP imports
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import * as strings from 'FeelingtaskWebPartStrings';
import Feelingtask from './components/Feelingtask';
import { IFeelingtaskProps } from './components/IFeelingtaskProps';

export interface IFeelingtaskWebPartProps {
  description: string;
}

export default class FeelingtaskWebPart extends BaseClientSideWebPart<IFeelingtaskWebPartProps> {
  private _sp: SPFI;

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // Initialize PnP SP
      this._sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));
    });
  }

  public render(): void {
    const element: React.ReactElement<IFeelingtaskProps> = React.createElement(
      Feelingtask,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        spInstance: this._sp // Pass the initialized SP instance
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}