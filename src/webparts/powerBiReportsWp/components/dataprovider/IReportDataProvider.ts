import { IReport } from '../models/IReport';
export interface IReportDataProvider {
    isValidList(string): Promise<boolean>;
    getItems(string): Promise<IReport[]>;
}
