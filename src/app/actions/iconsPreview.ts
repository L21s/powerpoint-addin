import {showErrorPopup} from "./errorPopup";
import {resetSearchInputAndDrawer} from "./searchDrawer";
import {iconsPreview} from "../taskpane";
import {FetchIconResponse} from "../shared/types";
import {downloadIconWith, fetchIcons, getDownloadPathForIconWith} from "../services/iconApiService";

const RECENT_ICONS_STORAGE_KEY = "recentIcons";
const MAX_NUMBER_OF_RECENT_ICONS = 30;

export async function fetchIconsForPreview(
    searchTerm: string,
    abortSignal: AbortSignal
): Promise<FetchIconResponse[]> {
  return searchTerm ? await fetchIcons(searchTerm, abortSignal) : loadRecentIconsFromLocalStorage();
}

export function addToPreview(icons: FetchIconResponse[]) {
  document.querySelectorAll("sl-skeleton").forEach((skeleton) => skeleton.remove());

  icons.forEach((icon) => {
    const buttonElement = document.createElement("sl-button") as HTMLButtonElement;
    buttonElement.id = icon.id;

    const iconElement = document.createElement("img");
    iconElement.src = icon.url;
    iconElement.slot = "prefix";

    iconsPreview.appendChild(buttonElement);
    buttonElement.appendChild(iconElement);
    buttonElement.onclick = (e) => insertSvgIcon(e, icon);
  });
}

async function insertSvgIcon(e: MouseEvent, icon: FetchIconResponse) {
  const button = e.target as HTMLButtonElement;
  button["loading"] = true;

  addIconToRecentIcons(icon);

  try {
    const path = await getDownloadPathForIconWith(icon.id);
    const svgText = await downloadIconWith(path).then((response) => response.text());

    Office.context.document.setSelectedDataAsync(
      svgText,
      { coercionType: Office.CoercionType.XmlSvg },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          const errorMessage = `Insert SVG failed: ${asyncResult.error.message}`;
          showErrorPopup(errorMessage);
        }
      }
    );
  } catch {
    throw new Error("Error inserting SVG icon: " + e);
  }

  button["loading"] = false;
  resetSearchInputAndDrawer();
}

function loadRecentIconsFromLocalStorage(): FetchIconResponse[] {
  const json = localStorage.getItem(RECENT_ICONS_STORAGE_KEY);
  return json ? JSON.parse(json) : [];
}

function saveRecentIconsToLocalStorage(icons: FetchIconResponse[]): void {
  localStorage.setItem(RECENT_ICONS_STORAGE_KEY, JSON.stringify(icons));
}

function addIconToRecentIcons(icon: FetchIconResponse): void {
  const recent = loadRecentIconsFromLocalStorage();
  const alreadyExists = recent.some((i) => i.id === icon.id);

  if (!alreadyExists) {
    const updated = [icon, ...recent].slice(0, MAX_NUMBER_OF_RECENT_ICONS);
    saveRecentIconsToLocalStorage(updated);
  }
}