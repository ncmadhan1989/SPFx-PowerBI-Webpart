import * as React from 'react';
import {
  DocumentCard, DocumentCardType,
  IDocumentCardStyles
} from 'office-ui-fabric-react/lib/DocumentCard';
import { IPowerBiReportsWpProps } from './IPowerBiReportsWpProps';
import ReportLists from './list/ReportLists';
import 'bootstrap/dist/css/bootstrap.min.css';

export default class PowerBiReportsWp extends React.Component<IPowerBiReportsWpProps, {}> {
  public render(): React.ReactElement<IPowerBiReportsWpProps> {
    const cardStyles: IDocumentCardStyles = {
      root: {
        width: '100%',
        maxWidth: 'none',
        minWidth: '300px',
        marginTop: '10px',
        backgroundColor: '#f3f2f1',
      }
    };

    return (
      <div className="container-fluid">
        <div className="row">
          <div className="col-lg-12 col-md-12 col-xs-12">
            <DocumentCard type={DocumentCardType.normal} styles={cardStyles}>
              <ReportLists
                siteurl={this.props.siteurl}
                listname={this.props.listname}
                iframeheight={this.props.iframeheight}
                reportsmenutitle={this.props.reportsmenutitle}
                webparttitle={this.props.webparttitle}
                openpropertypane={this.props.openpropertypane}
                >
              </ReportLists>
            </DocumentCard>
          </div>
        </div>
      </div>
    );
  }
}
