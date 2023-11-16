import { base64Image } from "./base64image";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("insert-image").onclick = run;
  }
});

export async function insertImage() {
    // Call Office.js to insert the image into the document.
    Office.context.document.setSelectedDataAsync(
      base64Image,
      {
        coercionType: Office.CoercionType.Image
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          //setMessage("Error: " + asyncResult.error.message);
        }
      }
    );
}

export async function run() {
  await PowerPoint.run(async (context) => {
  
    const shapes = context.presentation.getSelectedSlides().getItemAt(0).shapes
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 50;
    textbox.top = 50;
    textbox.height = 50;
    textbox.width = 150;
    textbox.name = "Square";
    textbox.fill.setSolidColor("yellow")
    textbox.textFrame.textRange.font.bold = true
    textbox.textFrame.textRange.font.name = "Arial"
    textbox.textFrame.textRange.font.size = 12
    textbox.textFrame.textRange.font.color = "#5A5A5A"
    textbox.lineFormat.visible = true
    textbox.lineFormat.color = "#000000"
    textbox.lineFormat.weight = 1.5
    console.log(textbox.lineFormat.toJSON)
    await context.sync();
  });
}