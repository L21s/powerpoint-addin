import {showErrorPopup} from "./errorPopup";
import {resetSearchInputAndDrawer} from "./searchDrawer";
import {iconsPreview} from "../taskpane";
import {FetchIconResponse} from "../shared/types";
import {downloadIconWith, fetchIcons, getDownloadPathForIconWith} from "../services/iconApiService";

export let recentIcons: FetchIconResponse[] = [];

export async function fetchIconsForPreview(
    searchTerm: string,
    abortSignal: AbortSignal
): Promise<FetchIconResponse[]> {
  return searchTerm ? await fetchIcons(searchTerm, abortSignal) : recentIcons;
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

function addIconToRecentIcons(icon: FetchIconResponse) {
  if (!recentIcons.includes(icon)) {
    recentIcons.unshift(icon);
    if (recentIcons.length > 12) recentIcons.pop();
  }
}

async function insertSvgIcon(e: MouseEvent, icon: FetchIconResponse) {
  const button = e.target as HTMLButtonElement;
  button["loading"] = true;
  addIconToRecentIcons(icon);
  const path = await getDownloadPathForIconWith(icon.id);
  const svgText = await downloadIconWith(path).then((response) => response.text());

  try {
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