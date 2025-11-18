import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneButton
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp/presets/all";
import * as strings from 'OrgchartWebPartStrings';
import Orgchart from './components/Orgchart';
import { IOrgchartProps } from './components/IOrgchartProps';

export interface IOrgchartWebPartProps {
  description: string;
  employeeCount: number;
  

}

export default class OrgchartWebPart extends BaseClientSideWebPart<IOrgchartWebPartProps> {
  public _isShowingAll: boolean = false;
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context as any
      });
    });
  }

  public render(): void {
    console.log('Rendering with employeeCount:', this.properties.employeeCount);
    const element: React.ReactElement<IOrgchartProps> = React.createElement(
      Orgchart,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        employeeCount: this.properties.employeeCount || 6,
        
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'employeeCount' && oldValue !== newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
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
                }),
                PropertyPaneSlider('employeeCount', {
                  label: 'Number of Employees to Display',
                  min: 1,
                  max: 30,
                  step: 1,
                  value: this.properties.employeeCount || 6,
                  showValue: true
                }),
                PropertyPaneButton('applyButton', {
                  text: 'Apply',
                  onClick: () => this.applyChanges()
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private applyChanges(): void {
    console.log('Applying changes with employeeCount:', this.properties.employeeCount);
    this.context.propertyPane.refresh();
    this.render();
  }
}