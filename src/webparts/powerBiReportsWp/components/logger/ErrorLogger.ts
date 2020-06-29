import { LogLevel, ILogListener, ILogEntry } from '@pnp/logging';
import { ReportDataProvider } from '../dataprovider/ReportDataProvider';
import { ILogItem } from './ILogItem';
import LogItem from './LogItem';

export default class ErrorLogger implements ILogListener {
    private _reportDataProvider: ReportDataProvider;
    private _applicationName: string;
    private _logListName: string;
    private _logWebUrl: string;
    private _currentUser: string;
    private _currentUserId: number;

    constructor(applicationName, logListName, logWebUrl, currentUser) {
        this._reportDataProvider = ReportDataProvider.getInstance();
        this._applicationName = applicationName;
        this._logListName = logListName;
        this._logWebUrl = logWebUrl;
        this._currentUser = currentUser;
    }

    private async getUserId(loginName: string): Promise<number> {
        let userId = 0;
        let userInfo = await this._reportDataProvider.getCurrentUserInfo(loginName);
        userId = userInfo.Id;
        return userId;
    }

    public async log(entry: ILogEntry): Promise<void> {
        try {
            let isValidList = await this._reportDataProvider.isValidList(this._logListName);
            if (isValidList) {
                this._currentUserId = await this.getUserId(this._currentUser);
                if (entry.level === LogLevel.Error) {
                    let newLogItem: ILogItem = new LogItem(
                        this._applicationName,
                        this._applicationName,
                        entry.data.CodeFileName,
                        entry.data.MethodName,
                        this._currentUserId,
                        new Date(),
                        entry.message,
                        entry.data.StackTrace
                    );
                    await this._reportDataProvider.addErrorLogItem(this._logListName, newLogItem);
                }
            }
        } catch (error) {
            console.log(`Error logging error to SharePoint list ${this._logListName} - ${error}`);
        }
        return;
    }

}