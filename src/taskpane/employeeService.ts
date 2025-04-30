const url = 'https://localhost:8443/employees/photos';
const imageCache = new Map<string, string>();
let recentEtag = "";

export async function fetchEmployeePhotos(): Promise<void> {

    //Todo: Append token
    const requestOptions = {
        method: "GET",
        headers: recentEtag ? { "If-None-Match": recentEtag } : {}
    };

    try {
        const result = await fetch(url, requestOptions);
        if (result.status === 304) {
            console.log("Images unchanged. Skipping update.");
            return;
        }

        const response = await result.json();
        updateLocalStorageImageCache(new Map(Object.entries(response)));

        const newEtag = result.headers.get("ETag");
        if (newEtag) {
           recentEtag = newEtag;
            console.log("New Etag set: ", newEtag);
        }
    } catch (e) {
        throw new Error("Error fetching employee photos: " + e);
    }
}

function getEmployeePhotos():string[] {
    return Array.from(imageCache.keys())
        .map(key => imageCache.get(key));
}

export function searchForImages(name: string) {
    if(!name) {
       return getEmployeePhotos();
   }

    return Array.from(imageCache.keys())
        .filter(key => key.includes(name))
        .map(key => imageCache.get(key));
}

export function tmpInsertSingleImage(base64String: string) {
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

function updateLocalStorageImageCache(map: Map<string, string>) {
    Array.from(map.keys()).forEach(key => imageCache.set(key, map.get(key)));
}