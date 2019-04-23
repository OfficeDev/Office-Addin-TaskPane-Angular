import { Component } from '@angular/core';
import * as OfficeHelpers from "@microsoft/office-js-helpers";
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
            return excel.run();
          case Office.HostType.OneNote:
            return onenote.run();
          case Office.HostType.Outlook:
            return outlook.run();
          case Office.HostType.PowerPoint:
            return powerpoint.run();
          case Office.HostType.Project:
            return project.run();
          case Office.HostType.Word:
            return word.run();
        }
      }
}