import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import * as powerpointComponent from "../../src/taskpane/app/powerpoint.app.component";

/* global describe, global, it */

const PowerPointMockData = {
  context: {
    document: {
      setSelectedDataAsync: function (data: string, options?) {
        this.data = data;
        this.options = options;
      },
    },
  },
  CoercionType: {
    Text: {},
  },
  onReady: async function () {},
};

describe("PowerPoint", function () {
  it("Run", async function () {
    const officeMock: OfficeMockObject = new OfficeMockObject(PowerPointMockData); // Mocking the common office-js namespace
    global.Office = officeMock as any;

    const powerpoint = new powerpointComponent.default();
    await powerpoint.run();

    assert.strictEqual(officeMock.context.document.data, "Hello World!");
  });
});
