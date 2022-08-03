import { Component } from "@angular/core";

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
