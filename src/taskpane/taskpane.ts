/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

import { base64Images } from "../../base64Image";
import { runPowerPoint } from "./powerPointUtil";
import { columnLineName, createColumns, createRows, rowLineName } from "./rowsColumns";
import { downloadIconWith, fetchIcons, getDownloadPathForIconWith } from "./iconDownloadUtils";
//import { storeEncryptionKey } from "./encryptionUtils";
import { FetchIconResponse, ShapeTypeKey } from "./types";
import { addColoredBackground, chooseNewColor, RGBAToHex } from "./iconUtils";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    /*
    let initials = <HTMLInputElement>document.getElementById("initials");
    initials.value = localStorage.getItem("initials");
    storeEncryptionKey();
    */

    initStickerButtons();
    initRowsAndColumnsButtons();

    // images for logo dropdown buttons
    (document.getElementById("currentWithText") as HTMLImageElement).src =
      "data:image/png;base64, " + base64Images["logoTextBlack"];
    (document.getElementById("currentWithoutText") as HTMLImageElement).src =
      "data:image/png;base64, " + base64Images["logoBlack"];

    changeAndInsertLogoImage();
    initDropdownPlaceholder();
    addIconSearch();
    openAndCloseDrawer();
    insertIconOnClickOnPreview();
    registerIconBackgroundTools();
  }
});

function openAndCloseDrawer() {
  const drawer = document.getElementById("search-drawer") as any;
  const wrapper = document.getElementById("wrapper") as HTMLElement;
  const openButtons = document.querySelectorAll(".search-open");
  const activeSearch = document.getElementById("active-search") as HTMLInputElement;

  openButtons.forEach((searchButton: HTMLInputElement) => {
    searchButton.onclick = () => {
      drawer.open = true;
      wrapper.style.overflow = "hidden";
      wrapper.scrollTo({
        top: 0,
        behavior: "smooth",
      });

      // carousel: make tabs switch depending on value searchButton.value ("icons" or "team")
      document.querySelectorAll(".drawerTabs").forEach((drawerTab: HTMLElement) => {
        drawerTab.style.display = "none";
      });
      document.getElementById(searchButton.value).style.display = "flex";
    };
  });

  document.getElementById("close-drawer").onclick = () => {
    drawer.open = false;
    wrapper.style.overflow = "scroll";
    activeSearch.value = "";
  };
}

function changeAndInsertLogoImage() {
  document.querySelectorAll(".logo-button, .image-button").forEach((button: HTMLElement) => {
    // on initialize: insert image for each button (selected + dropdown options)
    const initImageID = button.getAttribute("data-value");
    (document.getElementById(initImageID) as HTMLImageElement).src =
      "data:image/png;base64, " + base64Images[initImageID];

    // on click: change the current image inside the logo buttons, then insert
    button.onclick = () => {
      const imageID = button.getAttribute("data-value");
      const currentImage = document.getElementById(
        // which dropdown button was changed?
        imageID.includes("Text") ? "currentWithText" : "currentWithoutText"
      ) as HTMLImageElement;

      // on insert: changes the current image shown in the dropdown button
      currentImage.src = "data:image/png;base64, " + base64Images[imageID];
      currentImage.parentElement.setAttribute("data-value", imageID);

      // add a shadow filter if the current logo is white, otherwise remove it
      currentImage.style.filter = imageID.includes("White") ? "drop-shadow(0 0 1px #000000)" : "none";

      Office.context.document.setSelectedDataAsync(
        base64Images[imageID],
        { coercionType: Office.CoercionType.Image },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Action failed. Error: " + asyncResult.error.message);
          }
        }
      );
    };
  });
}

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
  document.querySelectorAll(".sticky-note").forEach((button) => {
    const color = window.getComputedStyle(button as HTMLElement).backgroundColor;
    (button as HTMLElement).onclick = () => insertSticker(color);
  });

  /*
  document.getElementById("save-initials").onclick = () =>
    localStorage.setItem("initials", (<HTMLInputElement>document.getElementById("initials")).value);
  */
}

async function deleteShapesByName(name: string) {
  await PowerPoint.run(async (context) => {
    const sheet = context.presentation.getSelectedSlides().getItemAt(0);
    sheet.load("shapes");
    await context.sync();
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

export async function insertSticker(color: string) {
  await runPowerPoint((powerPointContext) => {
    const today = new Date();
    const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
    const textBox = shapes.addTextBox(localStorage.getItem("initials") + ", " + today.toDateString() + "\n", {
      height: 50,
      left: 50,
      top: 50,
      width: 150
    });
    textBox.name = "Square";
    textBox.fill.setSolidColor(RGBAToHex(color));
    setStickerFontProperties(textBox);
  });
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
  const svgText = await downloadIconWith(path).then((response) => response.text());

  Office.context.document.setSelectedDataAsync(svgText, { coercionType: Office.CoercionType.XmlSvg }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      const errorMessage = `Insert SVG failed. Code: ${asyncResult.error.code}. Message: ${asyncResult.error.message}`;
      showErrorPopup(errorMessage);
    }
  });
}

function addIconSearch() {
  document.getElementById("icons").onclick = async () => {
    document.querySelectorAll("#icon-previews li").forEach((li) => li.remove());

    try {
      const searchTerm = (<HTMLInputElement>document.getElementById("icon-search-input")).value;
      const result = await fetchIcons(searchTerm);
      addIconPreviewWith(result);
    } catch (e) {
      const errorMessage = `Error executing icon search. Code: ${e.code}. Message: ${e.message}`;
      showErrorPopup(errorMessage);
    }
  };
}

function registerIconBackgroundTools() {
  document.querySelectorAll(".shape-option").forEach((button: HTMLElement) => {
    button.onclick = () => {
      addColoredBackground(button.getAttribute("data-value") as ShapeTypeKey);
    };
  });

  // when color-picker value is changed, update the selected color in the paint-bucket
  document.getElementById("background-color-picker").addEventListener("change", async () => {
    const colorSelect = document.getElementById("background-color-picker") as HTMLInputElement;
    chooseNewColor(colorSelect.value);
  });

  document.querySelectorAll(".fixed-color").forEach((button: HTMLElement) => {
    button.onclick = () => {
      chooseNewColor(RGBAToHex(button.style.backgroundColor));
    };
  });
}

function insertIconOnClickOnPreview() {
  document.getElementById("icon-previews").addEventListener("click", (event) => insertSvgIconOn(event), false);
}

function showErrorPopup(errorMessage: string) {
  const popup = document.getElementById("errorPopup");
  const popupText = document.getElementById("errorPopupText");
  const closeButton = document.getElementById("closePopupButton");

  if (popup && popupText && closeButton) {
    popupText.textContent = errorMessage;
    popup.style.display = "flex";
    closeButton.addEventListener("click", () => {
      popup.style.display = "none";
    });
  }
}

export function initDropdownPlaceholder() {
  /*
  const iconPreviewElement = document.getElementById("icon-previews");
  for (let i = 0; i < 15; i++) {
    const spanElement = document.createElement("span");
    const anchorElement = document.createElement("a");
    const listElement = document.createElement("li");
    iconPreviewElement.appendChild(listElement);
    listElement.appendChild(anchorElement);
    anchorElement.appendChild(spanElement);
  }
    */
}
