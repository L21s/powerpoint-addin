/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

import { base64Images } from "../../base64Image";
import * as M from "../../lib/materialize/js/materialize.min";
import { runPowerPoint } from "./powerPointUtil";
import { columnLineName, rowLineName, createColumns, createRows } from "./rowsColumns";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    M.AutoInit(document.body);

    let initials = <HTMLInputElement>document.getElementById("initials");
    initials.value = localStorage.getItem("initials");

    document.getElementById("fill-background").onclick = async () => {
      const colorPicker = <HTMLInputElement>document.getElementById("background-color");
      const selectedColor = colorPicker.value;
      await addBackground(selectedColor);
    };

    initStickerButtons();
    initRowsAndColumnsButtons();

    document.querySelectorAll(".logo-button").forEach((button) => {
      (button as HTMLElement).onclick = () => insertImageByBase64(button.getAttribute("data-value"));
    });

    document.getElementById("icons").onclick = async () => {
      document.querySelectorAll(".icon-results img").forEach((img) => img.remove());

      try {
        const searchTerm = (<HTMLInputElement>document.getElementById("icon-search-input")).value;
        const urls = await fetchIcons(searchTerm);
        urls.forEach((url) => {
          getImageElementWithSource(url);
        });
      } catch (e) {
        throw new Error("Error retrieving icon urls: " + e);
      }
    };
  }
});

function initRowsAndColumnsButtons() {
  document.getElementById("create-rows").onclick = () =>
    createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
  document.getElementById("delete-rows").onclick = () => deleteShapesByName(rowLineName);
  document.querySelectorAll(".row-button").forEach((button) => {
    (button as HTMLElement).onclick = () => {
      createRows(Number(button.getAttribute("data-value")));
    }
  });

  document.querySelectorAll(".column-button").forEach((button) => {
    (button as HTMLElement).onclick = () => createColumns(Number(button.getAttribute("data-value")));
  });
  document.getElementById("create-columns").onclick = () =>
    createColumns(+(<HTMLInputElement>document.getElementById("number-of-columns")).value);
  document.getElementById("delete-columns").onclick = () => deleteShapesByName(columnLineName);
}

function initStickerButtons() {
  document.querySelectorAll(".sticker-button").forEach((button) => {
    const color = window.getComputedStyle(button as HTMLElement).backgroundColor;
    (button as HTMLElement).onclick = () => insertSticker(color);
  });

  document.getElementById("save-initials").onclick = () =>
    localStorage.setItem("initials", (<HTMLInputElement>document.getElementById("initials")).value);
}

async function deleteShapesByName(name: string) {
  await PowerPoint.run(async (context) => {
    const sheet = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = sheet.shapes;

    shapes.load();
    await context.sync();

    shapes.items.forEach(function(shape) {
      if (shape.name == name) {
        shape.delete();
      }
    });
    await context.sync();
  });
}

function insertImageByBase64(base64Name: string) {
  Office.context.document.setSelectedDataAsync(
    base64Images[base64Name],
    { coercionType: Office.CoercionType.Image },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Action failed. Error: " + asyncResult.error.message);
      }
    }
  );
}


export async function insertSticker(color) {
  await runPowerPoint((powerPointContext) => {
    const today = new Date();
    const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
    const textBox = shapes.addTextBox(
      localStorage.getItem("initials") + ", " + today.toDateString() + "\n",
    { height: 50, left: 50, top: 50, width: 150 }
    );
    textBox.name = "Square";
    textBox.fill.setSolidColor(rgbToHex(color));
    setStickerFontProperties(textBox);
  });
}

function rgbToHex(rgb: String) {
  const regex = /(\d+),\s*(\d+),\s*(\d+)/;
  const matches = rgb.match(regex);
  function componentToHex(c: String) {
    const hex = Number(c).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  }

  return "#" + componentToHex(matches[1]) + componentToHex(matches[2]) + componentToHex(matches[3]);
}

function setStickerFontProperties(textbox: PowerPoint.Shape) {
  textbox.textFrame.textRange.font.bold = true;
  textbox.textFrame.textRange.font.name = "Arial";
  textbox.textFrame.textRange.font.size = 12;
  textbox.textFrame.textRange.font.color = "#5A5A5A";
  textbox.lineFormat.visible = true;
  textbox.lineFormat.color = "#000000";
  textbox.lineFormat.weight = 1.25;
}

export async function addBackground(backgroundColor?: string) {
  if (!backgroundColor) backgroundColor = "white";
  await runPowerPoint((powerPointContext) => {
    const selectedImage = powerPointContext.presentation.getSelectedShapes().getItemAt(0);
    selectedImage.fill.setSolidColor(backgroundColor);
  });
}

export async function fetchIcons(searchTerm: string): Promise<Array<string>> {
  const url = "https://hammerhead-app-fj5ps.ondigitalocean.app/icons?term=" + searchTerm;
  const requestHeaders = new Headers();
  requestHeaders.append("X-Freepik-API-Key", "FPSX6fb1f23cbea7497387b5e5b8eb8943de");
  const requestOptions = {
    method: "GET",
    headers: requestHeaders,
  };

  try {
    const result = await fetch(url, requestOptions);
    const response = await result.json();
    return response.data
      .filter((obj) => obj.author.name === "Smashicons" && obj.family.name === "Basic Miscellany Lineal")
      .map((obj) => obj.thumbnails[0].url)
      .slice(0, 50);
  } catch (e) {
    throw new Error("Error fetching icons: " + e);
  }
}

function getImageElementWithSource(source: string) {
  const iconUrlElement = document.getElementById("icon-urls");
  const imageElement = document.createElement("img");
  imageElement.src = source;
  imageElement.width = 50;
  imageElement.height = 50;
  iconUrlElement.appendChild(imageElement);
}
