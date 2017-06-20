import { Injectable } from '@angular/core';
import { IOfficeResult  } from './ioffice-result';

@Injectable()
export class ExcelService { 
    private workbook: Office.Document = Office.context.document;
    private bindingName: string = 'addinBinding';
    private namedItemName: string = "'Sheet1'!A1";
    private binding: Office.MatrixBinding;

    constructor() { }

    respondToDataChange(eventArgs: any) {
        // Do something
    }

    bindToWorkBook(): Promise<IOfficeResult> {
        return new Promise((resolve, reject) => {
            this.workbook.bindings.addFromNamedItemAsync(this.namedItemName, Office.BindingType.Matrix, { id: this.bindingName },
                (addBindingResult: Office.AsyncResult) => {
                    if (addBindingResult.status === Office.AsyncResultStatus.Failed) {
                        reject({
                            error: 'Unable to bind to workbook. Error: ' + addBindingResult.error.message
                        });
                    } else {
                        this.binding = addBindingResult.value;
                        resolve({
                            success: 'Created binding ' + this.bindingName + ' on ' + this.namedItemName
                        });
                    }
                });
        });
    }

    createHandlerOnA1(): Promise<Office.AsyncResultStatus> {
        return new Promise((resolve, reject) => {
            Office.select("bindings#addinBinding").addHandlerAsync(Office.EventType.BindingDataChanged, this.respondToDataChange, undefined, (result: Office.AsyncResult) => {
                if(result.status === Office.AsyncResultStatus.Failed) {
                    reject(result.status);
                } else {
                    resolve(result.status);
                }
            });
        })
    }

    createHandlerOnDoc(): Promise<Office.AsyncResultStatus> {
        return new Promise((resolve, reject) => {
            this.workbook.addHandlerAsync(Office.EventType.BindingDataChanged, this.respondToDataChange, undefined, (result: Office.AsyncResult) => {
                if(result.status === Office.AsyncResultStatus.Failed) {
                    reject(result.status);
                } else {
                    resolve(result.status);
                }
            });
        })
    }

}