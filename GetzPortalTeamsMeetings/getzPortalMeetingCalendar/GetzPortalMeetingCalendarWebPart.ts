import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'GetzPortalMeetingCalendarWebPartStrings';
import GetzPortalMeetingCalendar from './components/GetzPortalMeetingCalendar';
import { IGetzPortalMeetingCalendarProps } from './components/IGetzPortalMeetingCalendarProps';

export interface IGetzPortalMeetingCalendarWebPartProps {
  description: string;
}

export default class GetzPortalMeetingCalendarWebPart extends BaseClientSideWebPart<IGetzPortalMeetingCalendarWebPartProps> {

  
  public render(): void {
    const element: React.ReactElement<IGetzPortalMeetingCalendarProps> = React.createElement(
      GetzPortalMeetingCalendar,
      {
        context:this.context }
    );

    ReactDom.render(element, this.domElement);
  }

   protected onInit(): Promise<void> {
    const googleFontsLink = document.createElement('link');
    googleFontsLink.href = 'https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;700&display=swap';
    googleFontsLink.rel = 'stylesheet';
    document.head.appendChild(googleFontsLink);
    return super.onInit();
    
  }



  

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
