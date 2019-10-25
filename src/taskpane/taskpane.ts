/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, updateCharCount);
  }
});

export async function run() {
  // Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function(asyncResult) {
  //   if (asyncResult.status == Office.AsyncResultStatus.Failed) {
  //     write("Action failed. Error: " + asyncResult.error.message);
  //   } else {
  //     write("Selected data: " + asyncResult.value);
  //   }
  // });
}

export function updateCharCount({ document, type }: Office.DocumentSelectionChangedEventArgs) {
  document.getSelectedDataAsync<string>(Office.CoercionType.Text, function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      write("Action failed. Error: " + asyncResult.error.message);
    } else {
      write("Selected data: " + nhmrcCharCount(asyncResult.value));
    }
  });
}

// export async function run() {
//   return Word.run(async context => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
// }

// Function that writes to a div with id='message' on the page.
function write(message) {
  document.getElementById("message").innerText += message;
}

function nhmrcCharCount(text: string) {
  const charCount = text.length;
  const newlineCount = (text.match(/\r/g) || []).length;
  return charCount + newlineCount;
}
