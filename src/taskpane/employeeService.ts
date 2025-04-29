
//Todo: GitGub Pictures -> caching with e-Tag
import {employeeImage} from "./types";

const imageCache = new Map<string, employeeImage>;

export async function fetchEmployeePhotos(): Promise<void> {
    const url = 'https://localhost:8443/employees/photos';
    //Todo: Append token
    const requestOptions = { method: "GET" };

    try {
        const result = await fetch(url, requestOptions);
        const response = await result.json();
        const photosMap: Map<string, employeeImage> = new Map<string, employeeImage>(Object.entries(response));
        updateImageCache(photosMap);

        console.log("photosMap", photosMap);
        console.log(searchForImages('eichhorn'));
        const obj: employeeImage = searchForImages('eichhorn')[0] as employeeImage;
        console.log(obj);
        tmpInsertSingleImage(obj.base64value);

    } catch (e) {
        throw new Error("Error fetching employee photos: " + e);
    }
}

function updateImageCache(map: Map<string, employeeImage>) {
    map.forEach(((value, key) => {imageCache.set(key, value)}));
}

function searchForImages(name: string): employeeImage[] {
    return Array.from(imageCache.keys())
        .filter(key => key.includes(name))
        .map(key => imageCache.get(key));
}

function tmpInsertSingleImage(base64String: string) {
    Office.context.document.setSelectedDataAsync(
         base64String,
        { coercionType: Office.CoercionType.Image },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                const errorMessage = `Insert employee photo failed. Code: ${asyncResult.error.code}. Message: ${asyncResult.error.message}`;
                console.log(errorMessage)
            }
        });
}