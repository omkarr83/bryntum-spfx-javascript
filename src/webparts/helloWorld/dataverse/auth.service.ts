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
  if (response && response.accessToken) {
    const expiresIn = response.expiresOn
      ? Math.floor((response.expiresOn.getTime() - Date.now()) / 1000)
      : 3600; // default 1h if missing
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
 * Process redirect result from Microsoft login (hash/fragment).
 * Call this as early as possible when the webpart loads so we don't miss the token
 * and avoid a redirect loop. Idempotent.
 */
export async function processRedirectOnLoad(): Promise<void> {
  await getMsalInstance();
}

/**
 * Ensure user is signed in: try silent token first, then popup (to avoid redirect loop),
 * then redirect only if popup fails (e.g. blocked). Call on first load.
 */
export async function ensureLoginOrRedirect(): Promise<string | null> {
  let token = await getAccessToken();
  if (token) return token;

  const instance = await getMsalInstance();
  let accounts = instance.getAllAccounts();

  // If URL has auth response in hash, process it once more (in case handleRedirectPromise wasn't run yet)
  const hash = typeof window !== 'undefined' ? window.location.hash : '';
  if (hash && (hash.indexOf('access_token') !== -1 || hash.indexOf('code=') !== -1)) {
    await getMsalInstance(); // handleRedirectPromise is idempotent; may have been consumed
    token = getStoredToken();
    if (token) return token;
    accounts = instance.getAllAccounts();
  }

  if (accounts.length === 0) {
    // Prefer popup to avoid full-page redirect and hash-handling issues in SharePoint workbench
    try {
      const popupResult = await instance.loginPopup({ scopes: [dataverseScope] });
      if (popupResult && popupResult.accessToken && popupResult.expiresOn) {
        const expiresIn = Math.floor((popupResult.expiresOn.getTime() - Date.now()) / 1000);
        storeToken(popupResult.accessToken, expiresIn);
        return popupResult.accessToken;
      }
    } catch (popupErr: unknown) {
      const err = popupErr as { errorCode?: string; message?: string };
      const blocked = err.errorCode === 'popup_window_error' || err.errorCode === 'user_cancelled' ||
        (err.message && err.message.toLowerCase().indexOf('popup') !== -1);
      if (!blocked) throw popupErr;
      // Popup blocked or cancelled: fall back to redirect
      await loginRedirect();
      return null;
    }
    // No accounts and popup didn't run or didn't return token: use redirect
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
    try {
      const popupResponse = await instance.acquireTokenPopup({ scopes: [dataverseScope], account: accounts[0] });
      if (popupResponse.accessToken && popupResponse.expiresOn) {
        const expiresIn = Math.floor((popupResponse.expiresOn.getTime() - Date.now()) / 1000);
        storeToken(popupResponse.accessToken, expiresIn);
        return popupResponse.accessToken;
      }
    } catch {
      await instance.acquireTokenRedirect({ scopes: [dataverseScope], account: accounts[0] });
      return null;
    }
  }
  return null;
}

export function isAuthenticated(): boolean {
  return getStoredToken() !== null;
}
