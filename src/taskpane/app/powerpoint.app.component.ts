import { Component } from "@angular/core";

// images references in the manifest
/* eslint-disable no-unused-vars */
import icon16 from "../../../assets/icon-16.png";
import icon32 from "../../../assets/icon-32.png";
import icon64 from "../../../assets/icon-64.png";
import icon80 from "../../../assets/icon-80.png";
import icon128 from "../../../assets/icon-128.png";
/* eslint-enable no-unused-vars */

/* global Office, console */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      {
        coercionType: Office.CoercionType.Text,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  }
}
