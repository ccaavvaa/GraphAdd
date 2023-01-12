/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;
  if (!item) {
    return;
  }
  const itemId = item.itemId.split("/").join("-");
  const url = `http://localhost:5000/mail?id=${encodeURI(itemId)}`;
  if (Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
  } else {
    // eslint-disable-next-line no-undef
    window.open(url, "_blank");
  }
}
