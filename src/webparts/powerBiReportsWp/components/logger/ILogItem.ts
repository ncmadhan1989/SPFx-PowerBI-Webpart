export interface ILogItem{
    Title: string;
    ApplicationName: string;
    CodeFileName: string;
    MethodName: string;
    LoggedByUserId: number;
    LoggedOn: Date;
    ErrorMessage: string;
    StackTrace: string;
}