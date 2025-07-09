import {getRequestHeadersWithAuthorization} from "./authService";
import {FetchIconResponse} from "../shared/types";

const proxyBaseUrlIcons = `https://powerpoint-addin-ktor-pq9vk.ondigitalocean.app/icons`;
//const proxyBaseUrlIcons = `https://localhost:8443/icons`;

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

export async function getDownloadPathForIconWith(id: string) {
  const url = `${proxyBaseUrlIcons}/${id}/download?format=svg`;
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

export async function fetchIcons(searchTerm: string): Promise<Array<FetchIconResponse>> {
  const url = `${proxyBaseUrlIcons}?term=${searchTerm}`;
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
    throw new Error("Error fetching icons: " + e);
  }
}