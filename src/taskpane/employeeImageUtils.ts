import { FetchIconResponse, EmployeeName } from "./types";
import { showErrorPopup } from "./taskpane";
import { getAccessToken } from "../security/authService";

const proxyBaseUrl = `https://powerpoint-addin-ktor-pq9vk.ondigitalocean.app`;
export let recentNames: EmployeeName[] = [
  {
    id: "nachname-vorname",
    name: "Vorname Nachname",
  },
  {
    id: "nachname-vorname",
    name: "Vorname Nachname",
  },
];

export function addToTeamPreview(names: EmployeeName[]) {
  const teamPreviewElement = document.getElementById("team");

  names.forEach((name) => {
    const buttonElement = document.createElement("sl-menu-item") as HTMLButtonElement;
    buttonElement.id = name.id;
    buttonElement.innerText = name.name;

    teamPreviewElement.appendChild(buttonElement);
    buttonElement.onclick = (e) => insertEmployeeImage(e, name);
  });
}

function addNameToRecentNames(name: EmployeeName) {
  if (!recentNames.includes(name)) {
    recentNames.unshift(name);
    if (recentNames.length > 12) recentNames.pop();
  }
}

async function insertEmployeeImage(e: MouseEvent, name: EmployeeName) {
  const button = e.target as HTMLButtonElement;
  button["loading"] = true;
  addNameToRecentNames(name);

  /*
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
  */

  button["loading"] = false;
}

export async function fetchEmployeeImages(searchTerm: string): Promise<Array<EmployeeName>> {
  /*
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
  */
  return recentNames;
}
