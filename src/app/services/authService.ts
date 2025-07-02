import {getLoginRequest, getMsalApp, setInitials} from "../../security/authClient";
import {scopes} from "../../security/authConfig";

export async function getRequestHeadersWithAuthorization(): Promise<Headers> {
    const token = await getAccessToken();
    const requestHeaders =  new Headers();
    requestHeaders.append("Authorization", `Bearer ${token}`);
    return requestHeaders;

}

export async function getAccessToken(): Promise<string> {
    const msalApp = getMsalApp();
    const activeAccount = msalApp.getActiveAccount();
    const tokenRequest = {
        scopes: [scopes],
        account: activeAccount,
    };
    return  (await msalApp.acquireTokenSilent(tokenRequest)).accessToken;
}

export function getActiveAccount() {
    return getMsalApp().getActiveAccount();
}

export async function loginWithDialog() {
    try {
        const msalApp = getMsalApp();
        const loginRequest = getLoginRequest();

        const loginResponse = await msalApp.loginPopup(loginRequest);
        msalApp.setActiveAccount(loginResponse.account);
        setInitials();
        return loginResponse.account;
    } catch (error) {
        console.error("Login failed:", error);
        return null;
    }
}