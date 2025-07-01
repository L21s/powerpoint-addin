import { PublicClientApplication } from "@azure/msal-browser";
import {authUri, clientId, scopes} from "./authConfig";

const msalConfig = {
    auth: {
        clientId,
        authority: authUri,
        supportsNestedAppAuth: true,
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: true,
    },
};

const loginRequest = { scopes: [scopes] };

const msalApp = new PublicClientApplication(msalConfig);
msalApp.initialize();

export async function loginWithDialog() {
    try {
        const loginResponse = await msalApp.loginPopup(loginRequest);
        msalApp.setActiveAccount(loginResponse.account);
        setInitials();
        return loginResponse.account;
    } catch (error) {
        console.error("Login failed:", error);
        return null;
    }
}


export function getMsalApp() {
    return msalApp;
}

function setInitials() {
    const name = msalApp.getActiveAccount()?.name ?? "N/A";
    localStorage.setItem("initials", name);
}