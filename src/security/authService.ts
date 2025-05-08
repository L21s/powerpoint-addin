import { PublicClientApplication } from "@azure/msal-browser";

const clientId = "236e86e4-f190-48f7-be93-e794ed28382a";
const authUri = "https://login.microsoftonline.com/a40cd802-97bc-4645-88f7-89bff678a616/v2.0";
const scopes = "profile openid api://b4a241d6-4efe-497e-af56-8fc3a236f4d2/Freepik";

const msalConfig = {
    auth: {
        clientId: clientId,
        authority: authUri,
        supportsNestedAppAuth: true
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: true
    },
};

const loginRequest = {
    scopes: [scopes]
};

const msalApp = new PublicClientApplication(msalConfig);
msalApp.initialize();

export function loginWithDialog() {
    msalApp.loginPopup(loginRequest).then(loginResponse => {
        msalApp.setActiveAccount(loginResponse.account);
        setInitials();
    })
}

export async function getAccessToken(): Promise<string> {
    const activeAccount = msalApp.getActiveAccount();
    const tokenRequest = {
        scopes: [scopes],
        account: activeAccount,
    };
    return  (await msalApp.acquireTokenSilent(tokenRequest)).accessToken;
}

export function setInitials() {
    const initials = msalApp.getActiveAccount().name;
    localStorage.setItem("initials", initials);
}