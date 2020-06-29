import { IReport } from '../models/IReport';
import { IReportDataProvider } from './IReportDataProvider';
import { ILogItem } from '../logger/ILogItem';
import * as util from '../Util';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import "@pnp/sp/site-users";
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { LogLevel } from '@pnp/logging';
import { ISiteUserProps } from '@pnp/sp/site-users';
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

    public async getCurrentUserInfo(loginName): Promise<ISiteUserProps>{
        let user:ISiteUserProps = await (await sp.web.ensureUser(loginName)).data;
        return user;
    }

    public async addErrorLogItem(listTitle: string, item: ILogItem): Promise<void> {
        await sp.web.lists.getByTitle(listTitle)
            .items
            .add(item)
            .then((result) => {
                console.log('Error item added.');
            })
            .catch((error) => {
                console.log(`Error adding error log item: ${error}`);
            });
    }

    public async isValidList(listTitle: string): Promise<boolean> {
        return await sp.web.lists.getByTitle(listTitle)
            .select("ID")
            .get()
            .then((results) => {
                return true;
            })
            .catch((error) => {
                return false;
            });
    }

    public async getItems(listTitle: string): Promise<IReport[]> {
        let _reports: IReport[] = [];
        if (!listTitle)
            return;

        return await sp.web.lists.getByTitle(listTitle).items
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
            })
            .catch((error) => {
                util.logError("ReportDataProvider", "getItems", error.stack, LogLevel.Error, error.message);
                return [];
            });

    }

}