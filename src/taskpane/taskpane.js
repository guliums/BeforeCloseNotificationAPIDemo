/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run1").onclick = run1;
    document.getElementById("run2").onclick = run2;
    document.getElementById("run3").onclick = run3;
    document.getElementById("run4").onclick = run4;
  }
});

export async function run1() {
  try {
    await Office.addin.beforeDocumentCloseNotification.enable();
  } catch (error) {
    console.error(error);
  }
}

export async function run2() {
  try {
    await Office.addin.beforeDocumentCloseNotification.disable();
  } catch (error) {
    console.error(error);
  }
}

var removeEventHandler = null;
export async function run3() {
  try {
    if (removeEventHandler != null) {
      return;
    }

    removeEventHandler = await Office.addin.beforeDocumentCloseNotification.onCloseActionCancelled(async function () {
      console.log("event1 received");
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.values = [["onCloseActionCancelled event received!"]];
        await context.sync();
      });
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run4() {
  try {
    await removeEventHandler();
  } catch (error) {
    console.error(error);
  }
}
