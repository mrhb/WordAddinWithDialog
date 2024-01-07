/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

let dialog; // Declare dialog as global for use in later functions.
export async function run() {

  Office.context.ui.displayDialogAsync('https://localhost:3000/taskpanecopy.html', {height: 65, width: 65},

    function (asyncResult) {
      console.log('asyncResult');
      console.log(asyncResult);
      dialog = asyncResult.value; 
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (evt) => {
          // evt.message       
          document.getElementById("recievedMessage").value=evt.message ;
          console.log(evt.message       );
      })
    });

  // return Word.run(async (context) => {
  //   /**
  //    * Insert your Word code here
  //    */

  //   // insert a paragraph at the end of the document.
  //   const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

  //   // change the paragraph color to blue.
  //   paragraph.font.color = "blue";

  //   await context.sync();
  // });
}

export async function processDialogCallback(message) {
  console.log(message);
  return;
}
