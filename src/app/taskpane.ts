import { loginWithDialog } from "../security/authClient";
import {initializeStickyNotes} from "./init/stickyNotes";
import {initializeRowsColumns} from "./init/rowsColumns";
import {initializeSearchDrawer} from "./init/searchDrawer";
import {initializeImageBackgroundEditor} from "./init/imageBackgroundEditor";
import {initializeLogoDropdown} from "./init/logos";

// sticky notes
export const stickyNotes = document.querySelectorAll(".sticky-note");

// rows & columns
export const createRowsElement = document.getElementById("create-rows");
export const deleteRowsElement = document.getElementById("delete-rows");
export const rowButtons = document.querySelectorAll(".row-button");
export const createColumnsElement = document.getElementById("create-columns");
export const deleteColumnsElement = document.getElementById("delete-columns");
export const colButtons = document.querySelectorAll(".column-button");

// search & drawer
export const searchInput = document.getElementById("search-input") as HTMLInputElement;
export const drawer = document.getElementById("search-drawer") as HTMLElement;
export const activeDrawer = document.getElementById("active-drawer") as HTMLInputElement;
export const wrapper = document.getElementById("wrapper") as HTMLElement;

// image background editor
export const fixedColors = document.querySelectorAll(".fixed-color");
export const paintBucketColor = document.getElementById("paint-bucket-color");
export const shapeOptions = document.querySelectorAll(".shape-option");
export const backgroundColorPicker = document.getElementById("background-color-picker");

// logo dropdown
export const logoDropdownOptions = document.querySelectorAll(".logo-dropdown, .logo-dropdown-option");

// icons preview
export const iconsPreview = document.getElementById("icons");

//employees preview
export const employeesPreview = document.getElementById("names");

// popup
export const popup = document.querySelector("sl-alert") as any;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loginWithDialog();
    initializeTaskPane();
  }
});

export function initializeTaskPane() {
  initializeStickyNotes()
  initializeRowsColumns()
  initializeSearchDrawer()
  initializeImageBackgroundEditor()
  initializeLogoDropdown()
}