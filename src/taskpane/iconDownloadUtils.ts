import { FetchIconResponse } from "./types";
import { initDropdownPlaceholder } from "./taskpane";
import {getToken} from "../security/authService";

const baseUrl = `https://powerpoint-addin-ktor-pq9vk.ondigitalocean.app`;

export async function fetchIcons(searchTerm: string): Promise<Array<FetchIconResponse>> {
  const url = `${baseUrl}/icons?term=${searchTerm}`;
  const requestOptions = {
    method: "GET",
    headers: await getAuthorizedRequestHeaders(),
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
  const url = `${baseUrl}/icons/${id}/download?format=svg`;
  const requestOptions = {
    method: "GET",
    headers: await getAuthorizedRequestHeaders(),
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

async function getAuthorizedRequestHeaders(): Promise<Headers> {
  const token = await getToken();
  const requestHeaders = new Headers();
  requestHeaders.append("Authorization", `Bearer ${token}`);
  return requestHeaders;
}