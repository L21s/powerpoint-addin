/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

import { runPowerPoint } from "./powerPointUtil";
import { columnLineName, rowLineName, createColumns, createRows } from "./rowsColumns";
import { getDownloadPathForIconWith, downloadIconWith, fetchIcons } from "./iconDownloadUtils";
import { FetchIconResponse } from "./types";
import { loginWithDialog } from "../security/authService";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loginWithDialog();
    initStickerButtons();
    initRowsAndColumnsButtons();
    registerDrawerToggle();
    registerSearch();
    registerIconBackgroundTools();
    registerLogoImageInsert();
  }
});

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
    ? 'Search results for "' + searchTerm + '"'
    : "Recently used icons";
});

function registerSearch() {
  document.getElementById("search-input").addEventListener("sl-input", () => {
    (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "block";
    processInputChanges();
  });
}

function registerDrawerToggle() {
  const drawer = document.getElementById("search-drawer") as HTMLElement;
  const wrapper = document.getElementById("wrapper") as HTMLElement;
  const searchButtons = document.querySelectorAll(".search-open");

  searchButtons.forEach((searchButton: HTMLInputElement) => {
    searchButton.onclick = () => {
      drawer["open"] = true;
      wrapper.style.overflow = "hidden";
      wrapper.scrollTo({
        top: 0,
        behavior: "smooth",
      });
    };
  });

  document.getElementById("close-drawer").onclick = () => {
    drawer["open"] = false;
    wrapper.style.overflow = "scroll";

    (document.getElementById("search-input") as HTMLInputElement).value = "";
    (document.getElementById("active-search") as HTMLInputElement).value = "";
  };
}

function registerLogoImageInsert() {
  document.querySelectorAll(".logo-dropdown, .logo-dropdown-option").forEach((button: HTMLElement) => {
    button.onclick = async () => {
      const selectedImageSrc = button.getElementsByTagName("img")[0].src;
      const currentDropdownImage = document.getElementById(
        selectedImageSrc.includes("Text") ? "currentWithText" : "currentWithoutText"
      ) as HTMLImageElement;

      currentDropdownImage.src = selectedImageSrc;
      if (selectedImageSrc.includes("White")) {
        currentDropdownImage.classList.add("white-shadow");
      } else {
        currentDropdownImage.classList.remove("white-shadow");
      }

      Office.context.document.setSelectedDataAsync(
        // setSelectedDataAsync does not accept "data:image/png;base64," part of the base64 string -> remove it with split
        ((await getImageAsBase64(selectedImageSrc)) as string).split(",")[1],
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

async function getImageAsBase64(imageSrc: string) {
  const response = await fetch(imageSrc);
  const blob = await response.blob();

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(blob);

    reader.onload = () => resolve(reader.result as string);
    reader.onerror = (error) => reject(error);
  }) as Promise<string>;
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
    const color = button.getAttribute("data-color");
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
    textBox.fill.setSolidColor(color);
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
