import { Component } from "@angular/core";
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
     * Insert your Outlook code here
     */
  }
}
