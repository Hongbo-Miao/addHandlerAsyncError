import { Injectable } from '@angular/core';
import { IOfficeResult  } from './ioffice-result';
import { Subject }    from 'rxjs/Subject';

@Injectable()
export class ExcelService { 
    private workbook: Office.Document = Office.context.document;
    private bindingName: string = 'addinBinding';
    private namedItemName: string = "'Sheet1'!A1";
    private binding: Office.MatrixBinding;

    // Observable parameter change sources
    private inputParameterChange = new Subject<number>();

    inputParameterChanged$ = this.inputParameterChange.asObservable();

    constructor() { }

    changeInputParameter(eventArgs: any) {
        this.inputParameterChange.next(0);
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

    createHandlerOnA1(): Promise<IOfficeResult> {
        return new Promise((resolve, reject) => {
            this.workbook.bindings.getByIdAsync(this.bindingName, (result: Office.AsyncResult) => {
                if(result.status === Office.AsyncResultStatus.Failed) {
                    reject({
                        error: 'failed to get binding by id'
                    });
                } else {
                    result.value.addHandlerAsync(Office.EventType.BindingDataChanged, this.changeInputParameter, (handlerResult: Office.AsyncResult) => {
                        if(handlerResult.status === Office.AsyncResultStatus.Failed) {
                            reject({
                                error: 'failed to set a handler'
                            });
                        } else {
                            // Successful 
                            resolve({
                                success: 'successfully set handler'
                            });
                        }
                    })
                }
            })
        })
    }
}