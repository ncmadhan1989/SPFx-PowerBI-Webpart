import { IReport } from '../models/IReport';
import { IReportDataProvider } from './IReportDataProvider';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export class ReportDataProvider implements IReportDataProvider {

    private _siteurl: string = "";
    private _reportslistname: string = "";

    constructor(props) {
        this._siteurl = props.siteurl;
        this._reportslistname = props.listname;

        sp.setup({
            sp: {
                headers: {
                    Accept: "application/json;odata=verbose",
                },
                baseUrl: this._siteurl
            },
        });
    }

    public getItems(): Promise<IReport[]> {
        let _reports: IReport[] = [];
        if (this._siteurl && this._reportslistname) {
            return sp.web.lists.getByTitle(this._reportslistname).items
                .select("ID", "Title", "CategoryName", "ReportURL", "ReportName")
                .orderBy("ID", true)
                .get()
                .then((results: IReport[]) => {
                    results.forEach((res, index) => {
                        _reports.push({
                            ID: res.ID,
                            key: res.ID,
                            Title: res.Title,
                            ReportURL: res.ReportURL,
                            ReportName: res.ReportName,
                            CategoryName: res.CategoryName
                        });
                    });
                    return _reports;
                });
        }

    }

}