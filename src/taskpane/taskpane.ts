/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

import { base64Images } from "../../base64Image";
import * as M from "../../lib/materialize/js/materialize.min";
import { runPowerPoint } from "./powerPointUtil";
import { columnLineName, rowLineName, createColumns, createRows } from "./rowsColumns";
import { getDownloadPathForIconWith, downloadIconWith, fetchIcons } from "./iconDownloadUtils";
import { FetchIconResponse } from "./types";
import {loginWithDialog} from "../security/authService";


Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    M.AutoInit(document.body);
    loginWithDialog();
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

    initDropdownPlaceholder();
    addIconSearch();
    insertIconOnClickOnPreview();
  }
});

function initRowsAndColumnsButtons() {
  document.getElementById("create-rows").onclick = () =>
    createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
  document.getElementById("delete-rows").onclick = () => deleteShapesByName(rowLineName);
  document.querySelectorAll(".row-button").forEach((button) => {
    (button as HTMLElement).onclick = () => {
      createRows(Number(button.getAttribute("data-value")));
    };
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

function addIconPreviewWith(icons: FetchIconResponse[]) {
  for (let i = 0; i < icons.length; i += 5) {
    const iconPreviewElement = document.getElementById("icon-previews");
    const listElement = document.createElement("li");
    const anchorElement = document.createElement("a");
    iconPreviewElement.appendChild(listElement);
    listElement.appendChild(anchorElement);

    icons.slice(i, i + 5).forEach((icon) => {
      const iconPreviewElement = document.createElement("img");
      iconPreviewElement.id = icon.id;
      iconPreviewElement.src = icon.url;
      iconPreviewElement.width = 45;
      iconPreviewElement.height = 45;
      anchorElement.appendChild(iconPreviewElement);
    });
  }
}

async function insertSvgIconOn(event: any): Promise<void> {

  const path = await getDownloadPathForIconWith(event.target.id);
  const svgText = await downloadIconWith(path)
      .then((response) => response.text());

  Office.context.document.setSelectedDataAsync(
      svgText,
      { coercionType: Office.CoercionType.XmlSvg },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Insert SVG failed. Code: ${asyncResult.error.code}. Message: ${asyncResult.error.message}`);
        }
      }
  );
}


function addIconSearch() {
  document.getElementById("icons").onclick = async () => {
    document.querySelectorAll("#icon-previews li").forEach((li) => li.remove());

    try {
      const searchTerm = (<HTMLInputElement>document.getElementById("icon-search-input")).value;
      const result = await fetchIcons(searchTerm);
      addIconPreviewWith(result);
    } catch (e) {
      throw new Error("Error retrieving icon urls: " + e);
    }
  };
}

function insertIconOnClickOnPreview() {
  document.getElementById("icon-previews").addEventListener("click", (event) => insertSvgIconOn(event), false);
}

export function initDropdownPlaceholder() {
  const iconPreviewElement = document.getElementById("icon-previews");
  for (let i = 0; i < 15; i++) {
    const spanElement = document.createElement("span");
    const anchorElement = document.createElement("a");
    const listElement = document.createElement("li");
    iconPreviewElement.appendChild(listElement);
    listElement.appendChild(anchorElement);
    anchorElement.appendChild(spanElement);
  }
}
