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
      root: { width: '100%', maxWidth: 'none', minWidth: '800px' },
    };

    return (
      <div className="container-fluid">
        <div className="row">
          <div className="col-lg-12 col-md-12 col-xs-12">
            <DocumentCard type={DocumentCardType.normal} styles={cardStyles}>
              {
                (this.props.siteurl && this.props.listname) ?
                  <ReportLists siteurl={this.props.siteurl} listname={this.props.listname}></ReportLists>
                  :
                  <div className="alert alert-danger text-center" role="alert">
                    Please provide the Reports list name in the webpart properties pane.
                </div>
              }
            </DocumentCard>
          </div>
        </div>
      </div>
    );
  }
}
