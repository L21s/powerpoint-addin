import { FetchIconResponse } from "./types";
import { getDecryptedFreepikApiKey } from "./encryptionUtils";
import { initDropdownPlaceholder } from "./taskpane";

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
    showFetchIconsErrorInDropdown();
    initDropdownPlaceholder();
    throw new Error("Error fetching icons: " + e);
  }
}

export async function getDownloadPathForIconWith(id: string) {
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

function showFetchIconsErrorInDropdown() {
  const iconPreviewElement = document.getElementById("icon-previews");
  const spanElement = document.createElement("span");
  spanElement.innerText = "Error fetching icons";
  const anchorElement = document.createElement("a");
  const listElement = document.createElement("li");
  iconPreviewElement.appendChild(listElement);
  listElement.appendChild(anchorElement);
  anchorElement.appendChild(spanElement);
}
