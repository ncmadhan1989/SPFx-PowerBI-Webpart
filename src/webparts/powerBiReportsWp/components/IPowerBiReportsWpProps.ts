export interface IPowerBiReportsWpProps {
  description: string;
  siteurl: string;
  listname: string;
  iframeheight: number;
  reportsmenutitle: string;
  menuposition: string;
  webparttitle: string;
  errorloglist: string;
  paneltype: string;
  panelwidth: number;
  shownavigationpane: boolean;
  showfilterpane: boolean; 
  openpropertypane(): void;
}
