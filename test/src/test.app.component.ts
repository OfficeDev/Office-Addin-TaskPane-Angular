import { Component } from '@angular/core';
import * as excel from "./test.excel.app.component";
import * as word from "./test.word.app.component";
const template = require('./../../src/taskpane/app/app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async run() {
        switch (Office.context.host) {
            case Office.HostType.Excel:
                const excelComponent = new excel.default();
                return excelComponent.runTest();    
                const wordComponent = new word.default();
                return wordComponent.runTest();
        }
    }
}