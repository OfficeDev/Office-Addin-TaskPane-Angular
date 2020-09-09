import { Component } from "@angular/core";
// images references in the manifest
const icon16 = require("../../../assets/icon-16.png");
const icon32 = require("../../../assets/icon-32.png");
const icon80 = require("../../../assets/icon-80.png");
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    /**
     * Insert your OneNote code here
     */
  }
}
