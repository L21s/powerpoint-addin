import { PublicClientApplication } from "@azure/msal-browser";

const clientId = "236e86e4-f190-48f7-be93-e794ed28382a";
const redirectUri = "https://localhost:3000/taskpane.html";
const tenantId = "a40cd802-97bc-4645-88f7-89bff678a616";

const msalConfig = {
    auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/a40cd802-97bc-4645-88f7-89bff678a616/v2.0`,
        redirectUri: redirectUri,
        supportsNestedAppAuth: true,
        tenantId: tenantId
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false, //?
    },
};

const loginRequest = {
    scopes: ["profile openid api://b4a241d6-4efe-497e-af56-8fc3a236f4d2/Freepik"]
};

const msalApp = new PublicClientApplication(msalConfig);
msalApp.initialize();

export async function loginWithDialog() {
    await msalApp.loginPopup(loginRequest).then(loginResponse => {msalApp.setActiveAccount(loginResponse.account);})
    await setInitials();
}

export async function getToken(): Promise<string> {
    const activeAccount = msalApp.getActiveAccount();
    const tokenRequest = {
        scopes: ["profile openid api://b4a241d6-4efe-497e-af56-8fc3a236f4d2/Freepik"],
        account: activeAccount,
    };
    return  (await msalApp.acquireTokenSilent(tokenRequest)).accessToken;
}

export async function setInitials() {
    const initials = msalApp.getActiveAccount().name;
    localStorage.setItem("initials", initials);
}