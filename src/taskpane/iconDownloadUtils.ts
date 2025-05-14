import { FetchIconResponse } from "./types";
import { showErrorPopup } from "./taskpane";
import { getRequestHeadersWithAuthorization } from "../security/authService";

const proxyBaseUrl = `https://powerpoint-addin-ktor-pq9vk.ondigitalocean.app`;
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

function addIconToRecentIcons(icon: FetchIconResponse) {
  if (!recentIcons.includes(icon)) {
    recentIcons.unshift({
      id: icon.id,
      url: icon.url,
    });

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
          const errorMessage = `Insert SVG failed. Code: ${asyncResult.error.code}. Message: ${asyncResult.error.message}`;
          showErrorPopup(errorMessage);
        }
      }
    );
  } catch {
    throw new Error("Error inserting SVG icon: " + e);
  }

  button["loading"] = false;
}


export async function fetchIcons(searchTerm: string): Promise<Array<FetchIconResponse>> {
  const url = `${proxyBaseUrl}/icons?term=${searchTerm}`;
  const requestOptions = {
    method: "GET",
    headers: await getRequestHeadersWithAuthorization(),
  };

  try {
    const result = await fetch(url, requestOptions);
    const response = await result.json();
    return response.data
      .filter((obj: any) => obj.author.name === "Smashicons" && obj.family.name === "Basic Miscellany Lineal")
      .map((obj: any) => ({
        id: obj.id.toString(),
        url: obj.thumbnails[0].url,
      }))
      .slice(0, 50);
  } catch (e) {
    showErrorMessageInDrawer();
  }
}

function showErrorMessageInDrawer() {
  const iconPreviewElement = document.getElementById("icon-previews");
  const spanElement = document.createElement("div");
  spanElement.innerText = "Error fetching icons";
  iconPreviewElement.appendChild(spanElement);
}

export async function getDownloadPathForIconWith(id: string) {
  const url = `${proxyBaseUrl}/icons/${id}/download?format=svg`;
  const requestOptions = {
    method: "GET",
    headers: await getRequestHeadersWithAuthorization(),
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

export function debounce(func: Function) {
  let timer: NodeJS.Timeout;
  return (...args: any[]) => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      func.apply(this, args);
    }, 500);
  };
}