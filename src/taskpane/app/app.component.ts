import { Component } from '@angular/core';
import * as excel from "./excel.app.component";
import * as onenote from "./onenote.app.component";
import * as outlook from "./outlook.app.component";
import * as powerpoint from "./powerpoint.app.component";
import * as project from "./project.app.component";
import * as word from "./word.app.component";
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
            const excelTaskpane = new excel.default();
            return excelTaskpane.run();
          case Office.HostType.OneNote:
            const onenoteTaskpane = new onenote.default();
            return onenoteTaskpane.run();
          case Office.HostType.Outlook:
            const outlookTaskpane = new outlook.default();
            return outlookTaskpane.run();
          case Office.HostType.PowerPoint:
            const powerpointTaskpane = new powerpoint.default();
            return powerpointTaskpane.run();
          case Office.HostType.Project:
            const projectTaskpane = new project.default();
            return projectTaskpane.run();
          case Office.HostType.Word:
            const wordTaskpane = new word.default();
            return wordTaskpane.run();
        }
      }
}