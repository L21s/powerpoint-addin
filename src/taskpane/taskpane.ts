/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    let initials = <HTMLInputElement>document.getElementById("initials");
    initials.value = localStorage.getItem("initials");

    document.getElementById("fill-background").onclick = () => {
      const colorPicker = <HTMLInputElement>document.getElementById("background-color");
      const selectedColor = colorPicker.value;
      addBackground(selectedColor);
    };
    document.getElementById("remove-background").onclick = () => removeBackground();
    document.getElementById("yellow-sticker").onclick = () => insertSticker("yellow");
    document.getElementById("cyan-sticker").onclick = () => insertSticker("#00ffff");
    document.getElementById("save-initials").onclick = () =>
      localStorage.setItem("initials", (<HTMLInputElement>document.getElementById("initials")).value);
    document.getElementById("create-rows").onclick = () =>
        createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
    document.getElementById("create-columns").onclick = () =>
        createColumns(+(<HTMLInputElement>document.getElementById("number-of-columns")).value);
  }
});

export async function createRows(numberOfRows: number) {
  const lineDistance = 354 / numberOfRows
  let top = 126;

  for (let _i = 0; _i <= numberOfRows; _i++) {
    await runPowerPoint((powerPointContext) => {
      const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
      const line = shapes.addLine(PowerPoint.ConnectorType.straight);
      line.name = "StraightLine";
      line.left = 8;
      line.top = top;
      line.height = 0;
      line.width = 944;
    });

    top += lineDistance;
  }
}

export async function createColumns(numberOfColumns: number) {
  const lineDistance = 848 / numberOfColumns
  let left= 58;

  for (let _i = 0; _i <= numberOfColumns; _i++) {
    await runPowerPoint((powerPointContext) => {
      const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
      const line = shapes.addLine(PowerPoint.ConnectorType.straight);
      line.name = "StraightLine";
      line.left = left;
      line.top = 8;
      line.height = 524;
      line.width = 0;
    });

    left += lineDistance;
  }
}

function loadImageIntoLocalStorage(input?: HTMLInputElement) {
  if (!input) return;
  const file = input.files[0];
  const reader = new FileReader();
  reader.readAsDataURL(file);
  reader.onload = function () {
    const base64String = (reader.result as string).replace(new RegExp("^data.{0,}base64,"), "");
    localStorage.setItem("base64Image", base64String);
  };
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
    console.log(textbox.lineFormat.toJSON);
  });
}

export async function addBackground(backgroundColor?: string) {
  if (!backgroundColor) backgroundColor = "white";
  await runPowerPoint((powerPointContext) => {
    const selectedImage = powerPointContext.presentation.getSelectedShapes().getItemAt(0);
    selectedImage.fill.setSolidColor(backgroundColor);
  });
}

export async function removeBackground() {
  await runPowerPoint((powerPointContext) => {
    const selectedImage = powerPointContext.presentation.getSelectedShapes().getItemAt(0);
    selectedImage.fill.clear();
  });
}

export async function insertImageWithBackground(backgroundColor?: string) {
  if (!backgroundColor) backgroundColor = "white";
  const base64Image = localStorage.getItem("base64Image");
  await runPowerPoint((powerPointContext) => {
    Office.context.document.setSelectedDataAsync(
      base64Image,
      {
        coercionType: Office.CoercionType.Image,
      },
      async () => {
        const id = await getNewestShapeIdAsync();
        const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
        shapes.getItem(id).fill.setSolidColor(backgroundColor);
        await powerPointContext.sync();
      }
    );
  });
}

async function getNewestShapeIdAsync() {
  return await PowerPoint.run(async function (context) {
    const shapes = context.presentation.getSelectedSlides().getItemAt(0).shapes.load();
    await context.sync();
    const length = shapes.items.length;
    return shapes.items[length - 1].id;
  });
}

export async function runPowerPoint(updateFunction: (context: PowerPoint.RequestContext) => void) {
  await PowerPoint.run(async (context) => {
    updateFunction(context);
    await context.sync();
  });
}
