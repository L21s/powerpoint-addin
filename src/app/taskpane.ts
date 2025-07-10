import {initializeStickyNotesListener} from "./listener/stickyNotes";
import {initializeRowsColumnsListener} from "./listener/rowsColumns";
import {initializeSearchDrawerListener} from "./listener/searchDrawer";
import {initializeImageBackgroundEditorListener} from "./listener/backgroundFills";
import {initializeLogoDropdownListener} from "./listener/logos";
import {initializeBannerListener} from "./listener/banner";

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

// icons preview
export const iconsPreview = document.getElementById("icons");

// employees preview
export const employeesPreview = document.getElementById("names");

// background fills
export const fixedColors = document.querySelectorAll(".fixed-color");
export const paintBucket = document.getElementById("paint-bucket");
export const paintBucketColor = document.getElementById("paint-bucket-color");
export const shapeOptions = document.querySelectorAll(".shape-option");
export const backgroundColorPicker = document.getElementById("background-color-picker");
export const deleteBackground = document.getElementById("delete-background");

// logos
export const logoDropdownOptions = document.querySelectorAll(".logo-dropdown, .logo-dropdown-option");

// banner
export const addBannerButton = document.getElementById("add-banner") as HTMLButtonElement;
export const removeBannerButton = document.getElementById("remove-banner") as HTMLButtonElement;
export const bannerTextInput = document.getElementById("banner-text") as HTMLInputElement;
export const bannerTextColorInput = document.getElementById("banner-text-color") as HTMLInputElement;
export const bannerBackgroundColorInput = document.getElementById("banner-background-color") as HTMLInputElement;
export const bannerPositionSelect = document.getElementById("banner-position") as HTMLSelectElement;

// popup
export const popup = document.querySelector("sl-alert") as any;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    initializeTaskPaneListener();
  }
});

function initializeTaskPaneListener() {
  initializeStickyNotesListener()
  initializeRowsColumnsListener()
  initializeSearchDrawerListener()
  initializeImageBackgroundEditorListener()
  initializeLogoDropdownListener()
  initializeBannerListener()
}