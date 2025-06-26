import {getRequestHeadersWithAuthorization} from "./authService";

//const proxyBaseUrlEmployees = `https://powerpoint-addin-ktor-pq9vk.ondigitalocean.app/employees`;
const proxyBaseUrlEmployees = `https://localhost:8443/employees`;

let employeeNames: string[] = [];
let recentEtagNames: string = "";
const sharedImageCash: Map<string, string> = new Map<string, string>();

async function getRequestOptions(etag?: string): Promise<RequestInit> {
  const headers = await getRequestHeadersWithAuthorization();
  headers.append("If-None-Match", etag ? etag : "");

  return {
    method: "GET",
    headers: headers,
  };
}

export async function fetchEmployeeNames() {
  try {
    const result = await fetch(proxyBaseUrlEmployees + "/names", await getRequestOptions(recentEtagNames));
    if (result.status === 304) {
      console.log("Employee names unchanged. Skipping update.");
      return employeeNames;
    }
    employeeNames = await result.json();

    const newEtag = result.headers.get("ETag");
    if (newEtag) {
      recentEtagNames = newEtag;
    }
    return employeeNames;
  } catch (e) {
    throw new Error("Error fetching employee photos: " + e);
  }
}

export async function fetchEmployeeImage(name: string): Promise<string> {
  try {
    const result = await fetch(
      proxyBaseUrlEmployees + `/${name}`,
      await getRequestOptions(sharedImageCash.get(`etag_${name}`))
    );
    if (result.status === 304) {
      console.log("Employee image unchanged. Skipping update.");
      return sharedImageCash.get(`image_${name}`);
    }
    const response: string = await result.text();
    sharedImageCash.set(`image_${name}`, response);
    const newEtag = result.headers.get("ETag");
    if (newEtag) {
      sharedImageCash.set(`etag_${name}`, newEtag);
    }
    return response;
  } catch (e) {
    throw new Error("Error fetching employee image: " + e);
  }
}
