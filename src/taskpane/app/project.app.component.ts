import { Component } from "@angular/core";

// images references in the manifest
/* eslint-disable no-unused-vars */
import icon16 from "../../../assets/icon-16.png";
import icon32 from "../../../assets/icon-32.png";
import icon64 from "../../../assets/icon-64.png";
import icon80 from "../../../assets/icon-80.png";
import icon128 from "../../../assets/icon-128.png";
/* eslint-enable no-unused-vars */

/* global Office, console*/

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    try {
      // Get the GUID of the selected task
      Office.context.document.getSelectedTaskAsync((result) => {
        let taskGuid;
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          taskGuid = result.value;

          // Set the specified fields for the selected task.
          const targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
          const fieldValues = ["New task name", "Notes for the task."];

          // Set the field value. If the call is successful, set the next field.
          for (let index = 0; index < targetFields.length; index++) {
            Office.context.document.setTaskFieldAsync(taskGuid, targetFields[index], fieldValues[index], (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                index++;
              } else {
                console.log(result.error);
              }
            });
          }
        } else {
          console.log(result.error);
        }
      });
    } catch (error) {
      console.error(error);
    }
  }
}
