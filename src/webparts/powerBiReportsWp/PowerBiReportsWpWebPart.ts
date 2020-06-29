import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PowerBiReportsWpWebPartStrings';
import PowerBiReportsWp from './components/PowerBiReportsWp';
import { IPowerBiReportsWpProps } from './components/IPowerBiReportsWpProps';
import { ReportDataProvider } from './components/dataprovider/ReportDataProvider';
import ErrorLogger from './components/logger/ErrorLogger';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { Logger } from '@pnp/logging';

export interface IPowerBiReportsWpWebPartProps {
  description: string;
  listname: string;
  iframeheight: number;
  reportsmenutitle: string;
  webparttitle: string;
  errorloglist: string;
}

export default class PowerBiReportsWpWebPart extends BaseClientSideWebPart<IPowerBiReportsWpWebPartProps> {
  private _reportDataProvider: ReportDataProvider;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      this._reportDataProvider = ReportDataProvider.getInstance();
      sp.setup({
        spfxContext: this.context
      });
      this.registerLogging();
    });
  }

  public render(): void {
    const element: React.ReactElement<IPowerBiReportsWpProps> = React.createElement(
      PowerBiReportsWp,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        listname: this.properties.listname,
        iframeheight: this.properties.iframeheight,
        reportsmenutitle: this.properties.reportsmenutitle,
        webparttitle: this.properties.webparttitle,
        errorloglist: this.properties.errorloglist,
        openpropertypane: () => {
          this.context.propertyPane.open();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private registerLogging(): void {
    try {
      if (this.context && 
          this.context.pageContext && 
          this.properties.errorloglist) {
        let errorLoggerListener = new ErrorLogger(
          "PowerBIReportViewer",
          this.properties.errorloglist,
          this.context.pageContext.site.absoluteUrl,
          this.context.pageContext.user.loginName);
        Logger.subscribe(errorLoggerListener);
      }
    }
    catch (error) {
      console.log(`Error initializing error logger: ${error}`);
    }
    return;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this.registerLogging();
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
                  validateOnFocusOut: true,
                  onGetErrorMessage: this.validateListName.bind(this)
                }),
                PropertyPaneTextField('webparttitle', {
                  label: "Webpart title"
                }),
                PropertyPaneTextField('reportsmenutitle', {
                  label: "Menu title"
                }),
                PropertyPaneTextField('errorloglist', {
                  label: "Error list title",
                  validateOnFocusOut: true,
                  onGetErrorMessage: this.validateListName.bind(this)
                }),
                PropertyPaneSlider('iframeheight', {
                  label: 'Set IFrame height',
                  min: 300,
                  max: 1200,
                  value: 500,
                  showValue: true,
                  step: 10
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
      return "Provide the list title";
    }
    try {
      let isListExists = await this._reportDataProvider.isValidList(escape(value));
      if (isListExists)
        return '';

      return `List '${escape(value)}' doesn't exist in the current site`;
    }
    catch (error) {
      return error.message;
    }
  }

}
