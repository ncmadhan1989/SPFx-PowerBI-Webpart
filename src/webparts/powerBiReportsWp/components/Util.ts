import { LogLevel, ILogEntry, Logger } from "@pnp/logging";
import { ILogData } from './logger/ILogData';

export function logError(
    codefilename: string,
    methodname: string,
    stack: string,
    logLevel: LogLevel,
    error: any) {
    let data: ILogData = {
        CodeFileName: codefilename,
        MethodName: methodname,
        StackTrace: stack
    };
    let logEntry: ILogEntry = {
        message: `${error}`,
        level: logLevel,
        data: data
    };

    Logger.log(logEntry);
}
