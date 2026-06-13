import { AccountInfo, InteractionRequiredAuthError, PublicClientApplication } from "@azure/msal-browser";
import { loginScopes } from "./msalConfig";

export interface AuthSession {
  account: AccountInfo;
  accessToken: string;
}

export async function signIn(pca: PublicClientApplication): Promise<AccountInfo> {
  const response = await pca.loginPopup({ scopes: loginScopes, prompt: "select_account" });
  return response.account!;
}

export async function signOut(pca: PublicClientApplication, account: AccountInfo): Promise<void> {
  await pca.logoutPopup({ account });
}

export async function getAccessToken(
  pca: PublicClientApplication,
  account: AccountInfo
): Promise<string> {
  try {
    const token = await pca.acquireTokenSilent({ account, scopes: loginScopes });
    return token.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const token = await pca.acquireTokenPopup({ account, scopes: loginScopes });
      return token.accessToken;
    }

    throw error;
  }
}
