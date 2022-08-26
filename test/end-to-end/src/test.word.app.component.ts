import { Component } from '@angular/core';
import { pingTestServer, sendTestResults } from 'office-addin-test-helpers';
import * as testHelpers from './test-helpers';
import * as word from '../../../src/app/word.app.component';

/* global Office, Word */
const port: number = 4201;
let testValues: any = [];

@Component({
  selector: 'app-home',
  templateUrl: './app.component.html',
})
export default class AppComponent {
  welcomeMessage = 'Welcome';
  constructor() {
    Office.onReady(async () => {
      const testServerResponse: object = await pingTestServer(port);
      if (testServerResponse['status'] == 200) {
        this.runTest();
      }
    });
  }

  async runTest(): Promise<void> {
    try {
      // Execute taskpane code
      const wordComponent = new word.default();
      await wordComponent.run();
      await testHelpers.sleep(2000);

      // Get output of executed taskpane code
      Word.run(async (context) => {
        var firstParagraph = context.document.body.paragraphs.getFirst();
        firstParagraph.load('text');
        await context.sync();
        await testHelpers.sleep(2000);

        testHelpers.addTestResult(testValues, 'output-message', firstParagraph.text, 'Hello World');
        await sendTestResults(testValues, port);
        testValues.pop();
        Promise.resolve();
      });
    } catch {
      Promise.reject();
    }
  }
}
