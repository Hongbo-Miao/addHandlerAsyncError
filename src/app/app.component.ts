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
  <button O
  class="ms-Button ms-Button--primary" 
  type="submit"
  (click)="addHandler()"
  ><span class="ms-Button-label">Add handler to A1</span></button>
  <p>{{feedback | json}}</p>
  `,
})
export class AppComponent  { 
  name = 'Demo of addHandlerAsync Error';
  private feedback: string;
  private workbook: Office.Document = Office.context.document;
  private bindingName: string = 'addinBinding';
  private namedItemName: string = "'Sheet1'!A1";
  private binding: Office.MatrixBinding;

  constructor() {
      this.feedback = '';
  }

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

  createHandlerOnA1(): Promise<IOfficeResult> {
        return new Promise((resolve, reject) => {
        this.binding.addHandlerAsync(Office.EventType.BindingDataChanged, this.changeEvent.bind(this), (handlerResult: Office.AsyncResult) => {
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
        })
    }

  addHandler() {
    this.createHandlerOnA1()
    .then((result: any) => {
      this.feedback = result.success;
    }, (result: IOfficeResult) => {
      this.feedback = result.error;
    });
  }

  changeFeedback(message: string) {
    this.feedback = message
  }

  // Excel methods
  changeEvent(eventArgs: any) {
      
      Excel.run(async (context) => {
            const data = [
                [eventArgs.binding.id]
            ];
      const range = context.workbook.getSelectedRange()
      range.values = data;
      await context.sync();
    });
    // This won't run for some reason?
    this.changeFeedback('hello');
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


}
