import { IReport } from '../models/IReport';
export interface IReportDataProvider {
    getItems(): Promise<IReport[]>;
}
