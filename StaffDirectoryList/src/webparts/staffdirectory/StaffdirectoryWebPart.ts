import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {sp} from "@pnp/sp/presets/all"


import * as strings from 'StaffdirectoryWebPartStrings';
import Staffdirectory from './components/Staffdirectory';
import { IStaffdirectoryProps } from './components/IStaffdirectoryProps';

export interface IStaffdirectoryWebPartProps {
  description: string;
}

export default class StaffdirectoryWebPart extends BaseClientSideWebPart<IStaffdirectoryWebPartProps> {

  

  public render(): void {
    const element: React.ReactElement<IStaffdirectoryProps> = React.createElement(
      Staffdirectory,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl, // Ensure it's passing site URL
      }
        
    );

    ReactDom.render(element, this.domElement);
  }

  
  protected onInit(): Promise<void> {
    return super.onInit().then(_=>{
      sp.setup({
        spfxContext:this.context as any
      });
    })
    
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
