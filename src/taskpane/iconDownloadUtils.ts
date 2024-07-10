import { FetchIconResponse } from "./types";

export async function fetchIcons(searchTerm: string): Promise<Array<FetchIconResponse>> {
  const url = `https://hammerhead-app-fj5ps.ondigitalocean.app/icons?term=${searchTerm}&family-id=300&filters[shape]=outline&filters[color]=solid-black&filters[free_svg]=premium`;
  const requestHeaders = new Headers();
  requestHeaders.append("X-Freepik-API-Key", "FPSX6fb1f23cbea7497387b5e5b8eb8943de");
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
    throw new Error("Error fetching icons: " + e);
  }
}

export async function getDownloadPathForIconWith(id: string) {
  const url = `https://hammerhead-app-fj5ps.ondigitalocean.app/icons/${id}/download?format=png`;
  const requestHeaders = new Headers();
  requestHeaders.append("X-Freepik-API-Key", "FPSX6fb1f23cbea7497387b5e5b8eb8943de");
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
