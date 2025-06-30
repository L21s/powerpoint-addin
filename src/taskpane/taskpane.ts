import { loginWithDialog } from "../security/authClient";
import { insertSticker } from "./ui/stickyNotes";
import {columnLineName, createColumns, createRows, deleteShapesByName, rowLineName} from "./ui/rowsColumns";
import {
  closeDrawer,
  handleDrawerChange,
  handleSearchInput,
} from "./ui/searchDrawer";
import {ShapeTypeKey} from "./types";
import {addColoredBackground, chooseNewColor} from "./ui/imageBackgroundEditor";
import {handleLogoImageInsert} from "./ui/logoDropdown";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loginWithDialog();
    initializeUI();
  }
});

function initializeUI() {
  initializeStickyNotes()
  initializeRowsColumns()
  initializeSearchDrawer()
  initializeImageBackgroundEditor()
  initializeLogoDropdown()
}

// sticky notes
const stickyNotes = document.querySelectorAll(".sticky-note");

// rows & columns
const createRowsElement = document.getElementById("create-rows");
const deleteRowsElement = document.getElementById("delete-rows");
const rowButtons = document.querySelectorAll(".row-button");
const createColumnsElement = document.getElementById("create-columns");
const deleteColumnsElement = document.getElementById("delete-columns");
const colButtons = document.querySelectorAll(".column-button");

// search & drawer
export const searchInput = document.getElementById("search-input") as HTMLInputElement;
export const drawer = document.getElementById("search-drawer") as HTMLElement;
export const activeDrawer = document.getElementById("active-drawer") as HTMLInputElement;
export const wrapper = document.getElementById("wrapper") as HTMLElement;

// image background editor
export const fixedColors = document.querySelectorAll(".fixed-color");
export const paintBucketColor = document.getElementById("paint-bucket-color");
const shapeOptions = document.querySelectorAll(".shape-option");
const backgroundColorPicker = document.getElementById("background-color-picker");

// logo dropdown
const logoDropdownOptions = document.querySelectorAll(".logo-dropdown, .logo-dropdown-option");

// icons preview
export const iconsPreview = document.getElementById("icons");

//employees preview
export const employeesPreview = document.getElementById("names");

// popup
export const popup = document.querySelector("sl-alert") as any;

export function initializeStickyNotes() {
  stickyNotes.forEach((button) => {
    const color = button.getAttribute("data-color");
    (button as HTMLElement).onclick = () => insertSticker(color);
  });
}

export function initializeRowsColumns() {
  createRowsElement.onclick = () => createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
  deleteRowsElement.onclick = () => deleteShapesByName(rowLineName);
  rowButtons.forEach((button) => {
    (button as HTMLElement).onclick = () => createRows(Number(button.getAttribute("data-value")));
  });

  createColumnsElement.onclick = () =>
      createColumns(+(<HTMLInputElement>document.getElementById("number-of-columns")).value);
  deleteColumnsElement.onclick = () => deleteShapesByName(columnLineName);
  colButtons.forEach((button) => {
    (button as HTMLElement).onclick = () => createColumns(Number(button.getAttribute("data-value")));
  });
}

function initializeSearchDrawer(){
  initializeDrawer()
  initializeSearchInput()
}

function initializeDrawer() {
  activeDrawer.addEventListener("sl-change", async (e) => {
    await handleDrawerChange(e)
  });

  document.getElementById("close-drawer").onclick = () => {
    closeDrawer()
  };
}

function initializeSearchInput() {
  searchInput.addEventListener("sl-input", () => {
    handleSearchInput()
  });
}

function initializeImageBackgroundEditor() {
  shapeOptions.forEach((button: HTMLElement) => {
    button.onclick = () => addColoredBackground(button.getAttribute("data-value") as ShapeTypeKey);
  });

  backgroundColorPicker.addEventListener("change", async (e) => {
    chooseNewColor((e.target as HTMLInputElement).value);
  });

  fixedColors.forEach((button: HTMLElement) => {
    button.onclick = () => chooseNewColor(button.getAttribute("data-color"));
  });
}

function initializeLogoDropdown() {
  logoDropdownOptions.forEach((button: HTMLElement) => {
    button.onclick = async () => handleLogoImageInsert(button)
  });
}