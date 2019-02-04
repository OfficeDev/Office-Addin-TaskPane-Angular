import { Component } from '@angular/core';
import * as OfficeHelpers from '@microsoft/office-js-helpers';

@Component({
    selector: 'app-home',
    template: './app.component.html',
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async run() {
        try {
            await Excel.run(async context => {
                /**
                 * Insert your Excel code here
                 */
                const range = context.workbook.getSelectedRange();

            // Read the range address
            range.load('address');

            // Update the fill color
            range.format.fill.color = 'yellow';

            await context.sync();
            console.log(`The range address was ${range.address}.`);
            });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
}
