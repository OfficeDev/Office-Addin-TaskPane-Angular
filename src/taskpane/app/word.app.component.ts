import { Component } from "@angular/core";
// images references in the manifest
const icon16 = require("../../../assets/icon-16.png");
const icon32 = require("../../../assets/icon-32.png");
const icon80 = require("../../../assets/icon-80.png");
const template = require("./app.component.html");
/* global require, Word */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  }
}
