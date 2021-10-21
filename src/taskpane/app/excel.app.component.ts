import { Component } from "@angular/core";

// images references in the manifest
/* eslint-disable no-unused-vars */
import icon16 from "../../../assets/icon-16.png";
import icon32 from "../../../assets/icon-32.png";
import icon64 from "../../../assets/icon-64.png";
import icon80 from "../../../assets/icon-80.png";
import icon128 from "../../../assets/icon-128.png";
/* eslint-enable no-unused-vars */

/* global console, Excel */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    try {
      await Excel.run(async (context) => {
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
      console.error(error);
    }
  }
}
