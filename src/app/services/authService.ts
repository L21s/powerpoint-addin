import {getMsalApp} from "../../security/authClient";
import {scopes} from "../../security/authConfig";

export async function getRequestHeadersWithAuthorization(): Promise<Headers> {
    const token = await getAccessToken();
    const requestHeaders =  new Headers();
    requestHeaders.append("Authorization", `Bearer ${token}`);
    return requestHeaders;

}

async function getAccessToken(): Promise<string> {
    const msalApp = getMsalApp();
    const activeAccount = msalApp.getActiveAccount();
    const tokenRequest = {
        scopes: [scopes],
        account: activeAccount,
    };
    return  (await msalApp.acquireTokenSilent(tokenRequest)).accessToken;

}