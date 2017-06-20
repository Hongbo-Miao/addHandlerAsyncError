import { Component } from '@angular/core';
import { ExcelService } from './services/excel.service';
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
  <br>
  <button 
  class="ms-Button ms-Button--primary" 
  type="submit"
  (click)="addHandlerToDoc()"
  ><span class="ms-Button-label">Add handler to document overall</span></button>
  <p>{{feedback | json}}</p>
  `,
})
export class AppComponent  { 
  name = 'Demo of addHandlerAsync Error';
  feedback = '';

  constructor(private excelService: ExcelService) {}

  onBind(){
    this.excelService
    .bindToWorkBook()
    .then((result: IOfficeResult) => {
        this.feedback = result.success;
    }, (result: IOfficeResult) => {
        this.feedback = result.error;
    });
  }

  addHandler() {
    this.excelService.createHandlerOnA1()
    .then((result: Office.AsyncResultStatus) => {
      this.feedback = result.toString();
      //this.onResult(result);
    }, (result: IOfficeResult) => {
      console.log(result);
                this.feedback = result.toString();
              });
  }

  addHandlerToDoc() {
    this.excelService.createHandlerOnDoc()
    .then((result: Office.AsyncResultStatus) => {
      this.feedback = result.toString();
      //this.onResult(result);
    }, (result: IOfficeResult) => {
      console.log(result);
                this.feedback = result.toString();
              });
  }


}
