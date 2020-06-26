import { IReport } from '../models/IReport';
import { IReportDataProvider } from './IReportDataProvider';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
// Singleton class for CURD operation on sharepoint list
export class ReportDataProvider implements IReportDataProvider {
    private static _instance: ReportDataProvider;

    private constructor() {
        if (ReportDataProvider._instance != null) {
            throw new Error("Use getInstance() method to get the single instance of this class");
        }
    }

    public static getInstance(): ReportDataProvider {
        if (!ReportDataProvider._instance) {
            ReportDataProvider._instance = new ReportDataProvider();
        }
        return ReportDataProvider._instance;
    }

    public isValidList(listTitle: string): Promise<boolean> {
        return sp.web.lists.getByTitle(listTitle)
            .select("ID")
            .get()
            .then((results) => {
                return true;
            })
            .catch((error) => {
                return false;
            });
    }

    public getItems(listTitle: string): Promise<IReport[]> {
        let _reports: IReport[] = [];
        if (!listTitle)
            return;

        return sp.web.lists.getByTitle(listTitle).items
            .select("ID", "Title", "CategoryName", "SubCategory", "ReportURL", "ReportName")
            .filter("CategoryName ne null and SubCategory ne null")
            .orderBy("CategoryName,SubCategory", true)
            .get()
            .then((results: IReport[]) => {
                results.forEach((res, index) => {
                    _reports.push({
                        ID: res.ID,
                        key: res.ID,
                        Title: res.Title,
                        ReportURL: res.ReportURL,
                        ReportName: res.ReportName,
                        CategoryName: res.CategoryName,
                        SubCategory: res.SubCategory,
                    });
                });
                return _reports;
            });

    }

}