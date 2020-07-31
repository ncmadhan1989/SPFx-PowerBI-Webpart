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
  
### Deploying and Adding the Webpart to SharePoint online.
You can follow the microsoft documentation to deploy the SPFx webpart to a SharePoint page.
[Deploy your client-side web part to a SharePoint page](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page).





