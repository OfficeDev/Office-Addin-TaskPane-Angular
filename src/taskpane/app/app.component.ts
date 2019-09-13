import { Component } from "@angular/core";
import * as excel from "./excel.app.component";
import * as onenote from "./onenote.app.component";
import * as outlook from "./outlook.app.component";
import * as powerpoint from "./powerpoint.app.component";
import * as project from "./project.app.component";
import * as word from "./word.app.component";
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    /* global Office */
    switch (Office.context.host) {
      case Office.HostType.Excel: {
        const excelComponent = new excel.default();
        return excelComponent.run();
      }
      case Office.HostType.OneNote: {
        const onenoteComponent = new onenote.default();
        return onenoteComponent.run();
      }
      case Office.HostType.Outlook: {
        const outlookComponent = new outlook.default();
        return outlookComponent.run();
      }
      case Office.HostType.PowerPoint: {
        const powerpointComponent = new powerpoint.default();
        return powerpointComponent.run();
      }
      case Office.HostType.Project: {
        const projectComponent = new project.default();
        return projectComponent.run();
      }
      case Office.HostType.Word: {
        const wordComponent = new word.default();
        return wordComponent.run();
      }
    }
  }
}
