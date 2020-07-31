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



