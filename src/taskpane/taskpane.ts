/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

import { base64Images } from "../../base64Image";
import { runPowerPoint } from "./powerPointUtil";
import { columnLineName, createColumns, createRows, rowLineName } from "./rowsColumns";
import { debounce, fetchIcons, addToIconPreview, recentIcons } from "./iconDownloadUtils";
import { RGBAToHex, registerIconBackgroundTools } from "./iconUtils";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    initStickerButtons();
    initRowsAndColumnsButtons();
    openAndCloseDrawer();
    registerSearch();
    registerIconBackgroundTools();
    changeAndInsertLogoImage();
  }
});

function registerSearch() {
  const processInputChanges = debounce(async () => {
    const searchTerm = (<HTMLInputElement>document.getElementById("search-input")).value;

    try {
      document.getElementById("icon-previews").replaceChildren();
      let result = recentIcons;
      if (searchTerm) result = await fetchIcons(searchTerm);
      addToIconPreview(result);
    } catch (e) {
      const errorMessage = `Error executing icon search. Code: ${e.code}. Message: ${e.message}`;
      showErrorPopup(errorMessage);
    }

    (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "none";
    document.getElementById("search-result-title").innerText = searchTerm
      ? 'search results for "' + searchTerm + '"'
      : "Recently used icons";
  });

  // add debounced search to input
  document.getElementById("search-input").addEventListener("sl-input", () => {
    (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "block";
    processInputChanges();
  });
}

function openAndCloseDrawer() {
  const drawer = document.getElementById("search-drawer") as any;
  const wrapper = document.getElementById("wrapper") as HTMLElement;
  const searchButtons = document.querySelectorAll(".search-open");

  searchButtons.forEach((searchButton: HTMLInputElement) => {
    searchButton.onclick = () => {
      drawer.open = true;
      wrapper.style.overflow = "hidden";
      wrapper.scrollTo({
        top: 0,
        behavior: "smooth",
      });

      /** FOR LATER: make tabs switch depending on value searchButton.value ("icons" or "team") */
      /*
      // refocus input
      (document.getElementById("search-input") as HTMLInputElement).focus();

      document.querySelectorAll(".drawerTabs").forEach((drawerTab: HTMLElement) => {
        drawerTab.style.display = "none";
      });
      document.getElementById(searchButton.value).style.display = "flex";
      */
    };
  });

  document.getElementById("close-drawer").onclick = () => {
    drawer.open = false;
    wrapper.style.overflow = "scroll";
    // reset search input & radio button
    (document.getElementById("search-input") as HTMLInputElement).value = "";
    (document.getElementById("active-search") as HTMLInputElement).value = "";
  };
}

function changeAndInsertLogoImage() {
  // images for logo dropdown buttons
  (document.getElementById("currentWithText") as HTMLImageElement).src =
    "data:image/png;base64, " + base64Images["logoTextBlack"];
  (document.getElementById("currentWithoutText") as HTMLImageElement).src =
    "data:image/png;base64, " + base64Images["logoBlack"];

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
}

async function deleteShapesByName(name: string) {
  await PowerPoint.run(async (context) => {
    const sheet = context.presentation.getSelectedSlides().getItemAt(0);
    sheet.load("shapes");
    await context.sync();
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

export async function insertSticker(color: string) {
  await runPowerPoint((powerPointContext) => {
    const today = new Date();
    const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
    const textBox = shapes.addTextBox(localStorage.getItem("initials") + ", " + today.toDateString() + "\n", {
      height: 50,
      left: 50,
      top: 50,
      width: 150,
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

export function showErrorPopup(errorMessage: string) {
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
