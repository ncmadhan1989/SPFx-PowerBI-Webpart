import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
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
  menuposition: string;
  webparttitle: string;
  errorloglist: string;
  shownavigationpane: boolean;
  showfilterpane: boolean;
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
        menuposition: this.properties.menuposition,
        errorloglist: this.properties.errorloglist,
        shownavigationpane: this.properties.shownavigationpane,
        showfilterpane: this.properties.showfilterpane,
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
          displayGroupsAsAccordion: true,
          header: {
            description: "You can display multiple reports from Power BI embed url, which can be sourced from the SharePoint list as configured below."
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listname', {
                  label: 'Reports list title',
                  validateOnFocusOut: true,
                  onGetErrorMessage: this.validateListName.bind(this)
                }),
                PropertyPaneTextField('errorloglist', {
                  label: "Error list title",
                  validateOnFocusOut: true,
                  onGetErrorMessage: this.validateListName.bind(this)
                })
              ]
            },
            {
              groupName: "Webpart Configuration(s)",
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('webparttitle', {
                  label: "Webpart title"
                }),
                PropertyPaneTextField('reportsmenutitle', {
                  label: "Menu title"
                }),
                PropertyPaneChoiceGroup('menuposition', {
                  label: 'Menu Position (page referesh required)',
                  options: [{ key: 'left', text: 'Left' }, { key: 'right', text: 'Right' }]
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
                  label: "Webpart description"
                })
              ]
            },
            {
              groupName: "Report Configuration(s)",
              groupFields: [
                PropertyPaneToggle('shownavigationpane', {
                  label: 'Show Navigation Pane',
                  key: "shownavigationpane",
                  checked: false,
                  onText: "On",
                  offText: "Off"
                }),
                PropertyPaneToggle('showfilterpane', {
                  label: 'Show Filter Pane',
                  key: "showfilterpane",
                  checked: false,
                  onText: "On",
                  offText: "Off"
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
