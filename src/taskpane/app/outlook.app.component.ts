import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template,
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    /**
     * Insert your Outlook code here
     */
  }
}
