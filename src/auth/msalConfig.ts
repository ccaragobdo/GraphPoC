import { BrowserCacheLocation, Configuration } from "@azure/msal-browser";

const clientId = import.meta.env.VITE_MSAL_CLIENT_ID || "";
const tenantId = import.meta.env.VITE_MSAL_TENANT_ID || "common";

if (!clientId) {
  // eslint-disable-next-line no-console
  console.warn("Missing VITE_MSAL_CLIENT_ID. Sign-in will fail until it is configured.");
}

export const loginScopes = ["User.Read", "Files.Read.All", "Sites.Read.All"];

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: `${window.location.origin}${import.meta.env.BASE_URL}`
  },
  cache: {
    cacheLocation: BrowserCacheLocation.LocalStorage,
    storeAuthStateInCookie: false
  }
};
