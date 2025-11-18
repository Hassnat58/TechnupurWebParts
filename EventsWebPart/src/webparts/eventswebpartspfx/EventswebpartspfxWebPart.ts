import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  // PropertyPaneTextField,
  PropertyPaneSlider,
  // PropertyPaneButton
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// import * as strings from 'EventswebpartspfxWebPartStrings';
import Eventswebpartspfx from './components/Eventswebpartspfx';
import { IEventswebpartspfxProps } from './components/IEventswebpartspfxProps';

export interface IEventswebpartspfxWebPartProps {
  description: string;
  itemsToShow: number; // Added the property
}

export default class EventswebpartspfxWebPart extends BaseClientSideWebPart<IEventswebpartspfxWebPartProps> {
  //private _isDirty: boolean = false;

  public render(): void {
    const element: React.ReactElement<IEventswebpartspfxProps> = React.createElement(
      Eventswebpartspfx,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        itemsToShow: this.properties.itemsToShow || 5 // Default to 5 if not set
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context as any
      });
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Apply Button
  protected get disableReactivePropertyChanges(): boolean {
    return true; // Disable reactive property changes
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'itemsToShow' && oldValue !== newValue) {
      //this._isDirty = true; // Mark as dirty when itemsToShow changes
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          
          groups: [
            {
              
              groupFields: [
                PropertyPaneSlider('itemsToShow', {
                  label: 'Number of items to show',
                  min: 1,
                  max: 20,
                  value: this.properties.itemsToShow || 5, // Default value
                  showValue: true,
                  step: 1
                })

              ]
            }
          ]
        }
      ]
    };
  }

  // private _onApplyClick(): void {
  //   this._isDirty = false; // Reset the dirty flag
  //   this.render(); // Re-render the web part to apply changes
  // }
}