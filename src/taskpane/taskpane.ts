/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

const rowLineName = "RowLine";
const columnLineName = "ColumnLine";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    let initials = <HTMLInputElement>document.getElementById("initials");
    initials.value = localStorage.getItem("initials");

    document.getElementById("fill-background").onclick = () => {
      const colorPicker = <HTMLInputElement>document.getElementById("background-color");
      const selectedColor = colorPicker.value;
      addBackground(selectedColor);
    };
    document.getElementById("yellow-sticker").onclick = () => insertSticker("yellow");
    document.getElementById("cyan-sticker").onclick = () => insertSticker("#00ffff");
    document.getElementById("save-initials").onclick = () =>
      localStorage.setItem("initials", (<HTMLInputElement>document.getElementById("initials")).value);
    document.getElementById("create-rows").onclick = () =>
        createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
    document.getElementById("delete-rows").onclick = () => deleteShapesByName(rowLineName);
    document.getElementById("two-rows").onclick = () => createRows(2);
    document.getElementById("three-rows").onclick = () => createRows(3);
    document.getElementById("four-rows").onclick = () => createRows(4);
    document.getElementById("create-columns").onclick = () =>
        createColumns(+(<HTMLInputElement>document.getElementById("number-of-columns")).value);
    document.getElementById("delete-columns").onclick = () => deleteShapesByName(columnLineName);
    document.getElementById("two-columns").onclick = () => createColumns(2);
    document.getElementById("three-columns").onclick = () => createColumns(3);
    document.getElementById("four-columns").onclick = () => createColumns(4);
  }
});

async function deleteShapesByName(name: string) {
  await PowerPoint.run(async (context) => {
    const sheet = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = sheet.shapes;

    shapes.load();
    await context.sync();

    shapes.items.forEach(function (shape) {
      if (shape.name == name) {
        shape.delete();
      }
    });
    await context.sync();
  });
}

export async function createRows(numberOfRows: number) {
  const lineDistance = 354 / numberOfRows
  let top = 126;

  for (let _i = 0; _i <= numberOfRows; _i++) {
    await runPowerPoint((powerPointContext) => {
      const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
      const line = shapes.addLine(PowerPoint.ConnectorType.straight);
      line.name = rowLineName;
      line.left = 8;
      line.top = top;
      line.height = 0;
      line.width = 944;
      line.lineFormat.color = "#000000"
      line.lineFormat.weight = 0.5;
    });

    top += lineDistance;
  }
}

export async function createEmployeeImageShape() {
    await runPowerPoint( (powerPointContext) => {
        const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
        const circle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse);
        circle.lineFormat.color = "#E43292";
        circle.lineFormat.weight = 5;
    });
}

async function insertImageByURL(url: string) {
    const base64Image = linkToBase64(url);

    Office.context.document.setSelectedDataAsync(
        await base64Image,
        {
            coercionType: Office.CoercionType.Image
        },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Action failed. Error: " + asyncResult.error.message);
            }
        }
    );
}

async function linkToBase64(url: string) {
    const response = await fetch(url);
    const arrayBuffer = await response.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    let binaryString = '';
    uint8Array.forEach(value => binaryString += String.fromCharCode(value));
    return btoa(binaryString);
}

export async function createColumns(numberOfColumns: number) {
  const lineDistance = 848 / numberOfColumns
  let left= 58;

  for (let _i = 0; _i <= numberOfColumns; _i++) {
    await runPowerPoint((powerPointContext) => {
      const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
      const line = shapes.addLine(PowerPoint.ConnectorType.straight);
      line.name = columnLineName;
      line.left = left;
      line.top = 8;
      line.height = 524;
      line.width = 0;
      line.lineFormat.color = "#000000"
      line.lineFormat.weight = 0.5;
    });

    left += lineDistance;
  }
}

export async function insertSticker(color) {
  await runPowerPoint((powerPointContext) => {
    const today = new Date();
    const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
    const textbox = shapes.addTextBox(localStorage.getItem("initials") + ", " + today.toDateString() + "\n");
    textbox.left = 50;
    textbox.top = 50;
    textbox.height = 50;
    textbox.width = 150;
    textbox.name = "Square";
    textbox.fill.setSolidColor(color);
    textbox.textFrame.textRange.font.bold = true;
    textbox.textFrame.textRange.font.name = "Arial";
    textbox.textFrame.textRange.font.size = 12;
    textbox.textFrame.textRange.font.color = "#5A5A5A";
    textbox.lineFormat.visible = true;
    textbox.lineFormat.color = "#000000"
    textbox.lineFormat.weight = 1.25;
  });
}

export async function addBackground(backgroundColor?: string) {
  if (!backgroundColor) backgroundColor = "white";
  await runPowerPoint((powerPointContext) => {
    const selectedImage = powerPointContext.presentation.getSelectedShapes().getItemAt(0);
    selectedImage.fill.setSolidColor(backgroundColor);
  });
}

export async function runPowerPoint(updateFunction: (context: PowerPoint.RequestContext) => void) {
  await PowerPoint.run(async (context) => {
    updateFunction(context);
    await context.sync();
  });
}
