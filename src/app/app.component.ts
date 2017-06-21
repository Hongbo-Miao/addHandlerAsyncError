import { Component } from '@angular/core';
import { IOfficeResult } from './services/ioffice-result';

@Component({
  selector: 'my-app',
  template: 
  `<h1>{{name}}</h1>
  <button 
  class="ms-Button ms-Button--primary" 
  type="submit"
  (click)="onBind()"
  ><span class="ms-Button-label">Bind to A1</span></button>
  <button 
  class="ms-Button ms-Button--primary" 
  type="submit"
  (click)="addHandler()"
  ><span class="ms-Button-label">Add handler to A1</span></button>
  <p>{{feedback | json}}</p>
  `,
})
export class AppComponent  { 
  name = 'Demo of addHandlerAsync Error';
  feedback = '';
  private workbook: Office.Document = Office.context.document;
  private bindingName: string = 'addinBinding';
  private namedItemName: string = "'Sheet1'!A1";
  private binding: Office.MatrixBinding;

  constructor() {}

  ngOnInit() {
  }

  onBind(){
    this
    .bindToWorkBook()
    .then((result: IOfficeResult) => {
        this.feedback = result.success;
    }, (result: IOfficeResult) => {
        this.feedback = result.error;
    });
  }

  addHandler() {
    this.createHandlerOnA1()
    .then((result: any) => {
      this.feedback = result.success;
      //this.onResult(result);
    }, (result: IOfficeResult) => {
      console.log(result);
                this.feedback = result.error;
              });
  }

  changeFeedback() {
    this.feedback = 'hello'
  }

  // Excel methods
  changeEvent(eventArgs: any) {
      Excel.run(async (context) => {
            const data = [
                ["Hello World"]
            ];
      const range = context.workbook.getSelectedRange()
      range.values = data;
      await context.sync();
    });
    // Handler cannot run this
    // this.changeFeedback(); 
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
                    result.value.addHandlerAsync(Office.EventType.BindingDataChanged, this.changeEvent, (handlerResult: Office.AsyncResult) => {
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
