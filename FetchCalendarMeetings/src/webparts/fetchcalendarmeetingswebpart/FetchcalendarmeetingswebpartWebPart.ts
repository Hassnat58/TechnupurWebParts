import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';


import * as strings from 'FetchcalendarmeetingswebpartWebPartStrings';
import Fetchcalendarmeetingswebpart from './components/Fetchcalendarmeetingswebpart';
import { IFetchcalendarmeetingswebpartProps } from './components/IFetchcalendarmeetingswebpartProps';

export interface IFetchcalendarmeetingswebpartWebPartProps {
  description: string;
}

export default class FetchcalendarmeetingswebpartWebPart extends BaseClientSideWebPart<IFetchcalendarmeetingswebpartWebPartProps> {

  

  public render(): void {
    const element: React.ReactElement<IFetchcalendarmeetingswebpartProps> = React.createElement(
      Fetchcalendarmeetingswebpart,
      {
        // description: this.properties.description,
        context:this.context        
        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  //AJ CODE
  protected onInit(): Promise<void> {
    const googleFontsLink = document.createElement('link');
    googleFontsLink.href = 'https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;700&display=swap';
    googleFontsLink.rel = 'stylesheet';
    document.head.appendChild(googleFontsLink);
    return super.onInit();
    
  }

  //AJ CODE
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
