import { Component } from '@angular/core';

/* global Office, console */

@Component({
  selector: 'app-home',
  templateUrl: './app.component.html',
})
export default class AppComponent {
  welcomeMessage = 'Welcome';

  async run() {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      'Hello World!',
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
