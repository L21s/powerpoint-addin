import {getRequestHeadersWithAuthorization} from "../security/authService";

const proxyBaseUrlEmployees = `https://powerpoint-addin-ktor-pq9vk.ondigitalocean.app/employees`;
//const proxyBaseUrlEmployeesDEV = `https://localhost:8443/employees`;
let employeeNames: string[] =  [];
let recentEtagNames: string = "";


async function getRequestOptions(etag?: string): Promise<RequestInit> {

    const headers = await getRequestHeadersWithAuthorization()
    headers.append("If-None-Match", etag ? etag : "");

    return {
        method: "GET",
        headers: headers
    };
}

export async function fetchEmployeeNames() {
    try {
        const result = await fetch(proxyBaseUrlEmployees+'/names', await getRequestOptions(recentEtagNames));
        if (result.status === 304) {
            console.log("Employee names unchanged. Skipping update.");
            return employeeNames;
        }
        employeeNames = await result.json();

        const newEtag = result.headers.get("ETag");
        console.log(newEtag);
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
        const result = await fetch(proxyBaseUrlEmployees+`/${name}`, await getRequestOptions(localStorage.getItem(`etag_${name}`)));
        if (result.status === 304) {
            console.log("Employee image unchanged. Skipping update.");
            return localStorage.getItem(`image_${name}`);
        }
        const response: string = await result.text();
        localStorage.setItem(`image_${name}`, response);
        const newEtag = result.headers.get("ETag");
        if (newEtag) {
            localStorage.setItem(`etag_${name}`, newEtag);
        }
        console.log('fetchImage:', response );
        return response;
    } catch (e) {
        switch (e.name) {
            case "QuotaExceededError": {
                console.warn('Local Storage Limit exceeded: clearing cache.');
                await fetchEmployeeImage(name);
                return;
            }
            default: {
                throw new Error("Error fetching employee image: " + e);
            }
        }
    }
}