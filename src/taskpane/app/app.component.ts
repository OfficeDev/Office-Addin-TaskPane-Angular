import { Component } from '@angular/core';
import * as OfficeHelpers from "@microsoft/office-js-helpers";
const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async run() {
        switch (Office.context.host) {
          case Office.HostType.Excel:
            return this.runExcel();
          case Office.HostType.OneNote:
            return this.runOneNote();
          case Office.HostType.Outlook:
            return this.runOutlook();
          case Office.HostType.PowerPoint:
            return this.runPowerPoint();
          case Office.HostType.Project:
            return this.runProject();
          case Office.HostType.Word:
            return this.runWord();
        }
      }
      
      async runExcel() {
        try {
          await Excel.run(async context => {
            /**
             * Insert your Excel code here
             */
            const range = context.workbook.getSelectedRange();
      
            // Read the range address
            range.load("address");
      
            // Update the fill color
            range.format.fill.color = "yellow";
      
            await context.sync();
            console.log(`The range address was ${range.address}.`);
          });
        } catch (error) {
          OfficeHelpers.UI.notify(error);
          OfficeHelpers.Utilities.log(error);
        }
      }
      
      async runOneNote() {
        /**
         * Insert your OneNote code here
         */
      }
      
      
      async runOutlook() {
        /**
         * Insert your Outlook code here
         */
      }
      
      async runPowerPoint() {
        /**
         * Insert your PowerPoint code here
         */
        Office.context.document.setSelectedDataAsync("Hello World!",
          {
            coercionType: Office.CoercionType.Text
          },
          result => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error(result.error.message);
            }
          }
        );
      }
      
      async runProject() {
        /**
         * Insert your Outlook code here
         */
      }
      
      async runWord() {
        return Word.run(async context => {
          /**
           * Insert your Word code here
           */
          const range = context.document.getSelection();
      
          // Read the range text
          range.load("text");
      
          // Update font color
          range.font.color = "red";
      
          await context.sync();
          console.log(`The selected text was ${range.text}.`);
        });
    }
}