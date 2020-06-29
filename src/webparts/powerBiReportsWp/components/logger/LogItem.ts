import { ILogItem } from './ILogItem';

export default class LogItem implements ILogItem {
    public Title: string;
    public ApplicationName: string;
    public CodeFileName: string;
    public MethodName: string;
    public LoggedByUserId: number;
    public LoggedOn: Date;
    public ErrorMessage: string;
    public StackTrace: string;

    constructor(_title,
        _applicationName,
        _codeFileName,
        _methodName,
        _loggedByUser,
        _loggedOn,
        _errorMessage,
        _stackTrace) {
        this.Title = _title;
        this.ApplicationName = _applicationName;
        this.CodeFileName = _codeFileName;
        this.MethodName = _methodName;
        this.LoggedByUserId = _loggedByUser;
        this.LoggedOn = _loggedOn;
        this.ErrorMessage = _errorMessage;
        this.StackTrace = _stackTrace;
    }

}