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
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';

export interface IPowerBiReportsWpWebPartProps {
  description: string;
  listname: string;
}

export default class PowerBiReportsWpWebPart extends BaseClientSideWebPart<IPowerBiReportsWpWebPartProps> {


  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IPowerBiReportsWpProps> = React.createElement(
      PowerBiReportsWp,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
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
                PropertyPaneTextField('listname', {
                  label: 'Reports list name',
                  onGetErrorMessage: this.validateListName.bind(this)
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

  private async validateListName(value: string): Promise<string> {
    if (value === null || value.length === 0) {
      return "Provide the list name";
    }
    try {
      return sp.web.lists.getByTitle(escape(value))
        .select("ID")
        .get()
        .then((result) => {
          return "";
        })
        .catch((error) => {
          return `List '${escape(value)}' doesn't exist in the current site`;
        });

    } catch (error) {
      return error.message;
    }
    return '';
  }

}
