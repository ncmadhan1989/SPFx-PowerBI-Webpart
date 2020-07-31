# SPFx PowerBI Reports Viewer Webpart
SharePoint Framework (SPFx) webpart to display the Multiple PowerBI reports in SPFx webpart, the reports link are sourced from the SharePoint list (ReportsConfig), which can be provided from the webpart properties.

## Webpart Properties
The following properties have been added to configure the webpart.<br/>
  - List configurations
    - Reports list title: The sharepoint custom list where we store all the Reports information required for the webpart.
    - Error log list: The sharepoint custom list used to log any error in the webpart.
  - Webpart configurations
    - Webpart title: Title of the webpart to display on the top panel.
    - Menu title: Title of the reports menu lists.
    - Menu position: Reports menu can be display on left or right side (left menu or right menu).
    - Set IFrame height: You can configure the height of the IFrame where the Power BI report get loaded.
    - Webpart description: Description of the webpart.
  - PowerBI Report Configurations
    - Show Navigation Pane: Show/Hide the page navigation of the PowerBI reports.
    - Show Filter Pane: Show/Hide the filters pane of the PowerBI report. 
    
## Prerequisites
### Set up your SharePoint Framework development environment
You can follow the microsoft documentation to setup development environment for SharePoint Framework.
[Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment).

### Running the application locally
Clone or download the project in your local repository - git clone <https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart> <br/>
Open the cloned/downloaded project folder and type the following command to install all the npm packages and dependencies mentioned in the package.json file. <br/>
```
npm install
```
To run the application on workbench and serve the localhost resources, type the following command <br/>
  - Open the SharePoint online workbench site <https://<tenant>/<site url>/_layouts/15/workbench.aspx
  - Add the webpart name 'PowerBI Reports Viewer Webpart'.
  - Configure the webpart by providing the properties value.
```
gulp serve --nobrowser
```
  
### Deploying the Webpart to SharePoint online.
You can follow the microsoft documentation to deploy the SPFx webpart to a SharePoint page.
[Deploy your client-side web part to a SharePoint page](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page).

To deploy the webpart in SharePoint online, run the command, this will create "power-bi-spfx-webparts.sppkg" file under sharepoint/solution folder and then you can upload this file to the AppCatalog site and deploy.
```
gulp build
```
```
gulp bundle --ship
```
```
gulp package-solution --ship
```
### Adding Webpart to SharePoint
You can add webpart in a sharepoint page in two ways.
  - add webpart on a modern page.
  - add webpart in a full-width column layout.
  - add webpart as a Single Part App Page - you can learn more about this here [Single Part App Page](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/single-part-app-pages?tabs=pnpposh)
 
### Screenshots of webpart on SharePoint page.

![full-width page](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi1.png)

![configure](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi2.png)

![properties](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi3.png)

![right menu](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi4.png)

![right menu](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi4.1.png)

![report with menu](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi5.png)

![report without menu](https://github.com/Vikas-Salvi/SPFx-PowerBI-Webpart/blob/master/sharepoint/assets/powerbi6.png)
  
## Resources
- **Setup Microsoft 365 tenant** - [Set up your Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- **Setup development environment** - [Set up your SharePoint Framework development environment](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
- **FluentUI React - [FluentUI controls](https://developer.microsoft.com/en-us/fluentui#/controls/web)
- **PnP JS library** - [PnP JS getting started](https://pnp.github.io/pnpjs/), [PnP JS list items operations](https://pnp.github.io/pnpjs/sp/items/), [Use (PnPJS) library with SharePoint Framework web parts](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/use-sp-pnp-js-with-spfx-web-parts)
