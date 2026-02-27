import { PublicClientApplication, SilentRequest, AuthenticationResult } from '@azure/msal-browser';
import { msalConfig, dataverseScope } from './dataverseConfig';

const TOKEN_STORAGE_KEY = 'dataverse_access_token';
const TOKEN_EXPIRY_KEY = 'dataverse_token_expiry';

let msalInstance: PublicClientApplication | null = null;

function storeToken(token: string, expiresIn?: number): void {
  try {
    localStorage.setItem(TOKEN_STORAGE_KEY, token);
    if (expiresIn) {
      const expiryTime = Date.now() + (expiresIn * 1000);
      localStorage.setItem(TOKEN_EXPIRY_KEY, expiryTime.toString());
    }
  } catch (e) {
    console.error('[Auth] Error storing token', e);
  }
}

function getStoredToken(): string | null {
  try {
    const token = localStorage.getItem(TOKEN_STORAGE_KEY);
    const expiryTime = localStorage.getItem(TOKEN_EXPIRY_KEY);
    if (!token) return null;
    if (expiryTime) {
      const expiry = parseInt(expiryTime, 10);
      if (Date.now() >= expiry) {
        localStorage.removeItem(TOKEN_STORAGE_KEY);
        localStorage.removeItem(TOKEN_EXPIRY_KEY);
        return null;
      }
    }
    return token;
  } catch {
    return null;
  }
}

export async function getMsalInstance(): Promise<PublicClientApplication> {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(msalConfig);
    await msalInstance.initialize();
  } else if (!msalInstance.getConfiguration()) {
    await msalInstance.initialize();
  }
  const response = await msalInstance.handleRedirectPromise();
  if (response && response.accessToken && response.expiresOn) {
    const expiresIn = Math.floor((response.expiresOn.getTime() - Date.now()) / 1000);
    storeToken(response.accessToken, expiresIn);
  }
  return msalInstance;
}

export async function getAccessToken(): Promise<string | null> {
  try {
    const instance = await getMsalInstance();
    const accounts = instance.getAllAccounts();
    if (accounts.length === 0) {
      return null;
    }
    const account = accounts[0];

    const stored = getStoredToken();
    if (stored) {
      const expiryTime = localStorage.getItem(TOKEN_EXPIRY_KEY);
      if (expiryTime) {
        const expiry = parseInt(expiryTime, 10);
        const timeUntilExpiry = expiry - Date.now();
        if (timeUntilExpiry > 5 * 60 * 1000) return stored;
      } else {
        return stored;
      }
    }

    const request: SilentRequest = {
      scopes: [dataverseScope],
      account
    };
    try {
      const response: AuthenticationResult = await instance.acquireTokenSilent(request);
      if (response.accessToken && response.expiresOn) {
        const expiresIn = Math.floor((response.expiresOn.getTime() - Date.now()) / 1000);
        storeToken(response.accessToken, expiresIn);
      }
      return response.accessToken;
    } catch (silentError: unknown) {
      const err = silentError as { errorCode?: string };
      if (err.errorCode === 'interaction_required' || err.errorCode === 'consent_required' ||
          err.errorCode === 'login_required' || err.errorCode === 'token_expired') {
        const popupResponse = await instance.acquireTokenPopup({ scopes: [dataverseScope], account });
        if (popupResponse.accessToken && popupResponse.expiresOn) {
          const expiresIn = Math.floor((popupResponse.expiresOn.getTime() - Date.now()) / 1000);
          storeToken(popupResponse.accessToken, expiresIn);
        }
        return popupResponse.accessToken;
      }
      throw silentError;
    }
  } catch (e) {
    console.error('[Auth] getAccessToken error', e);
    return null;
  }
}

export async function login(): Promise<string | null> {
  const instance = await getMsalInstance();
  const response = await instance.loginPopup({ scopes: [dataverseScope] });
  if (response.accessToken && response.expiresOn) {
    const expiresIn = Math.floor((response.expiresOn.getTime() - Date.now()) / 1000);
    storeToken(response.accessToken, expiresIn);
    return response.accessToken;
  }
  return null;
}

/**
 * Redirect to Microsoft login (SPA flow). Use when running inside SharePoint
 * to avoid AADSTS9002326; the app must be registered as "Single-page application"
 * in Azure AD with redirect URI matching the SharePoint site.
 */
export async function loginRedirect(): Promise<void> {
  const instance = await getMsalInstance();
  await instance.loginRedirect({ scopes: [dataverseScope] });
}

/**
 * Ensure user is signed in: try silent token first, then redirect to login if needed.
 * Call on first load; no login button required. After redirect, page reloads with token.
 */
export async function ensureLoginOrRedirect(): Promise<string | null> {
  const token = await getAccessToken();
  if (token) return token;
  const instance = await getMsalInstance();
  const accounts = instance.getAllAccounts();
  if (accounts.length === 0) {
    await loginRedirect();
    return null;
  }
  try {
    const t = await instance.acquireTokenSilent({ scopes: [dataverseScope], account: accounts[0] });
    if (t.accessToken && t.expiresOn) {
      const expiresIn = Math.floor((t.expiresOn.getTime() - Date.now()) / 1000);
      storeToken(t.accessToken, expiresIn);
      return t.accessToken;
    }
  } catch {
    await instance.acquireTokenRedirect({ scopes: [dataverseScope], account: accounts[0] });
    return null;
  }
  return null;
}

export function isAuthenticated(): boolean {
  return getStoredToken() !== null;
}
