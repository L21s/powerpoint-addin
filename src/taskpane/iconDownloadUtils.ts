import { FetchIconResponse } from "./types";
import { getDecryptedFreepikApiKey } from "./encryptionUtils";
import { showErrorPopup } from "./taskpane";

export let recentIcons = [];

export function addToIconPreview(icons: FetchIconResponse[]) {
  const iconPreviewElement = document.getElementById("icon-previews");

  icons.forEach((icon) => {
    const buttonElement = document.createElement("sl-button") as HTMLButtonElement;
    buttonElement.id = icon.id;

    const iconElement = document.createElement("img");
    iconElement.src = icon.url;
    iconElement.slot = "prefix";

    iconPreviewElement.appendChild(buttonElement);
    buttonElement.appendChild(iconElement);
    buttonElement.onclick = (e) => insertSvgIcon(e, icon);
  });
}

async function insertSvgIcon(e: MouseEvent, icon: FetchIconResponse) {
  const button = e.target as HTMLButtonElement;

  // show loading spinner on button while SVG is loaded into slide
  button["loading"] = true;

  const path = await getDownloadPathForIconWith(icon.id);
  const svgText = await downloadIconWith(path).then((response) => response.text());

  // add the icon to list of recently used icons (only if not already added, max. 12)
  if (!recentIcons.includes(icon)) {
    recentIcons.unshift({
      id: icon.id,
      url: icon.url,
    });

    if (recentIcons.length > 12) recentIcons.pop();
  }

  // insert SVG
  try {
    Office.context.document.setSelectedDataAsync(
      svgText,
      { coercionType: Office.CoercionType.XmlSvg },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          const errorMessage = `Insert SVG failed. Code: ${asyncResult.error.code}. Message: ${asyncResult.error.message}`;
          showErrorPopup(errorMessage);
        }
      }
    );
  } catch {
    throw new Error("Error inserting SVG icon: " + e);
  }

  // reset spinner / button state at the end
  button["loading"] = false;
}

export async function fetchIcons(searchTerm: string): Promise<Array<FetchIconResponse>> {
  const url = `https://hammerhead-app-fj5ps.ondigitalocean.app/icons?term=${searchTerm}&family-id=300&filters[shape]=outline&filters[color]=solid-black&filters[free_svg]=premium`;
  const requestHeaders = new Headers();
  requestHeaders.append("X-Freepik-API-Key", getDecryptedFreepikApiKey());
  const requestOptions = {
    method: "GET",
    headers: requestHeaders,
  };

  try {
    const result = await fetch(url, requestOptions);
    const response = await result.json();
    return response.data
      .filter((obj) => obj.author.name === "Smashicons" && obj.family.name === "Basic Miscellany Lineal")
      .map((obj) => ({
        id: obj.id.toString(),
        url: obj.thumbnails[0].url,
      }))
      .slice(0, 50);
  } catch (e) {
    const iconPreviewElement = document.getElementById("icon-previews");
    const spanElement = document.createElement("div");
    spanElement.innerText = "Error fetching icons";
    iconPreviewElement.appendChild(spanElement);

    throw new Error("Error fetching icons: " + e);
  }
}

async function getDownloadPathForIconWith(id: string) {
  const url = `https://hammerhead-app-fj5ps.ondigitalocean.app/icons/${id}/download?format=svg`;
  const requestHeaders = new Headers();
  requestHeaders.append("X-Freepik-API-Key", getDecryptedFreepikApiKey());
  const requestOptions = {
    method: "GET",
    headers: requestHeaders,
  };

  try {
    const result = await fetch(url, requestOptions);
    const response = await result.json();
    return response.data.url;
  } catch (e) {
    throw new Error("Error getting download url: " + e);
  }
}

export async function downloadIconWith(url: string) {
  const requestOptions = {
    method: "GET",
  };

  try {
    return await fetch(url, requestOptions);
  } catch (e) {
    throw new Error("Error downloading icon: " + e);
  }
}

export function debounce(func, timeout = 500) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      func.apply(this, args);
    }, timeout);
  };
}
