import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PowerBiReportsWpWebPartStrings';
import PowerBiReportsWp from './components/PowerBiReportsWp';
import { IPowerBiReportsWpProps } from './components/IPowerBiReportsWpProps';

export interface IPowerBiReportsWpWebPartProps {
  description: string;
  siteurl: string;
  listname: string;
}

export default class PowerBiReportsWpWebPart extends BaseClientSideWebPart<IPowerBiReportsWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPowerBiReportsWpProps> = React.createElement(
      PowerBiReportsWp,
      {
        description: this.properties.description,
        siteurl: this.properties.siteurl,
        listname: this.properties.listname,
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
                PropertyPaneTextField('siteurl', {
                  label: 'Base Site Url'
                }),
                PropertyPaneTextField('listname', {
                  label: 'Reports list name',
                }),
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
