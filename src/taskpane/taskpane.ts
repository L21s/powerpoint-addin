/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, PowerPoint */

import { runPowerPoint } from "./powerPointUtil";
import { columnLineName, rowLineName, createColumns, createRows } from "./rowsColumns";
import { addToIconPreview, debounce, fetchIcons, recentIcons, showMessageInDrawer } from "./iconDownloadUtils";
import { addToTeamPreview, filterEmployeeNames } from "./employeeImageUtils";
import { loginWithDialog } from "../security/authService";
import { registerIconBackgroundTools } from "./iconUtils";
import { EmployeeName, FetchIconResponse } from "./types";

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

const processInputChanges = debounce(async (activeDrawerTab: string) => {
  const searchTerm = (<HTMLInputElement>document.getElementById("search-input")).value;
  let result: FetchIconResponse[] | EmployeeName[] = [];

  try {
    if (activeDrawerTab === "icons") {
      result = searchTerm ? await fetchIcons(searchTerm) : recentIcons;
      addToIconPreview(result as FetchIconResponse[]);
    } else if (activeDrawerTab === "team") {
      result = await filterEmployeeNames(searchTerm);
      addToTeamPreview(result as EmployeeName[]);
    }
  } catch (e) {
    showErrorPopup("Could not fetch any icons/images: " + e.message);
  }

  (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "none";
  const searchResultTitle = document.getElementById(activeDrawerTab + "-search-title");
  if (searchTerm) searchResultTitle.innerText = 'Search results for "' + searchTerm + '"';
  else if (activeDrawerTab === "icons") searchResultTitle.innerText = "Recently used icons";
  else if (activeDrawerTab === "team") searchResultTitle.innerText = "All employees";

  if (result.length === 0) {
    if (activeDrawerTab === "team") showMessageInDrawer("No names fitting this search query");
    if (activeDrawerTab === "icons" && !searchTerm) showMessageInDrawer("No recent icons yet");
  }
});

function registerSearch() {
  document.getElementById("search-input").addEventListener("sl-input", () => {
    refreshSearchResults();
  });
}

function registerDrawerToggle() {
  const drawer = document.getElementById("search-drawer") as HTMLElement;
  const wrapper = document.getElementById("wrapper") as HTMLElement;

  document.getElementById("active-drawer").addEventListener("sl-change", (e) => {
    const activeDrawerTab = (e.target as HTMLInputElement).value;
    refreshSearchResults();

    const searchInput = document.getElementById("search-input");
    searchInput.focus();
    if (activeDrawerTab === "icons") searchInput.setAttribute("placeholder", "search icons...");
    if (activeDrawerTab === "team") searchInput.setAttribute("placeholder", "search names...");

    drawer["open"] = true;
    wrapper.style.overflow = "hidden";
    wrapper.scrollTo({
      top: 0,
      behavior: "smooth",
    });

    const tabs = document.querySelector("sl-split-panel") as any;
    if (activeDrawerTab === "icons") tabs.position = 100;
    else tabs.position = 0;
  });

  document.getElementById("close-drawer").onclick = () => {
    drawer["open"] = false;
    wrapper.style.overflow = "scroll";

    (document.getElementById("search-input") as HTMLInputElement).value = "";
    (document.getElementById("active-drawer") as HTMLInputElement).value = "";
  };
}

function refreshSearchResults() {
  const activeDrawerTab = (<HTMLInputElement>document.getElementById("active-drawer")).value;

  if (activeDrawerTab) {
    document.getElementById(activeDrawerTab).replaceChildren();
    (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "block";

    for (let i = 0; i < 12; i++) {
      const skeleton = document.createElement("sl-skeleton");
      skeleton.classList.add(activeDrawerTab);
      skeleton.setAttribute("effect", "pulse");
      document.getElementById(activeDrawerTab).appendChild(skeleton);
    }

    processInputChanges(activeDrawerTab);
  }
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
