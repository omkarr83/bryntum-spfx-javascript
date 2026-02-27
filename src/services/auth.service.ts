import { PublicClientApplication, AccountInfo, SilentRequest, AuthenticationResult } from '@azure/msal-browser';
import { msalConfig, dataverseRequest } from '../config/msalConfig';

let msalInstance: PublicClientApplication | null = null;

// localStorage keys
const TOKEN_STORAGE_KEY = 'dataverse_access_token';
const TOKEN_EXPIRY_KEY = 'dataverse_token_expiry';

/**
 * Store access token in localStorage
 */
function storeToken(token: string, expiresIn?: number): void {
    try {
        localStorage.setItem(TOKEN_STORAGE_KEY, token);
        if (expiresIn) {
            const expiryTime = Date.now() + (expiresIn * 1000);
            localStorage.setItem(TOKEN_EXPIRY_KEY, expiryTime.toString());
        }
    } catch (error) {
        console.error('Error storing token in localStorage:', error);
    }
}

/**
 * Get access token from localStorage
 */
function getStoredToken(): string | null {
    try {
        const token = localStorage.getItem(TOKEN_STORAGE_KEY);
        const expiryTime = localStorage.getItem(TOKEN_EXPIRY_KEY);
        
        if (!token) {
            return null;
        }
        
        // Check if token is expired
        if (expiryTime) {
            const expiry = parseInt(expiryTime, 10);
            if (Date.now() >= expiry) {
                // Token expired, remove it
                clearStoredToken();
                return null;
            }
        }
        
        return token;
    } catch (error) {
        console.error('Error getting token from localStorage:', error);
        return null;
    }
}

/**
 * Clear stored token from localStorage
 */
function clearStoredToken(): void {
    try {
        localStorage.removeItem(TOKEN_STORAGE_KEY);
        localStorage.removeItem(TOKEN_EXPIRY_KEY);
    } catch (error) {
        console.error('Error clearing token from localStorage:', error);
    }
}

/**
 * Set MSAL instance (called from AuthProvider after initialization)
 */
export function setMsalInstance(instance: PublicClientApplication): void {
    msalInstance = instance;
}

/**
 * Initialize MSAL instance
 */
export async function initializeMsal(): Promise<PublicClientApplication> {
    if (!msalInstance) {
        msalInstance = new PublicClientApplication(msalConfig);
        await msalInstance.initialize();
    } else if (!msalInstance.getConfiguration()) {
        // If instance exists but not initialized, initialize it
        await msalInstance.initialize();
    }
    return msalInstance;
}

/**
 * Get MSAL instance
 */
export async function getMsalInstance(): Promise<PublicClientApplication> {
    if (!msalInstance) {
        console.log('MSAL instance not found, initializing...');
        return await initializeMsal();
    }
    // Ensure it's initialized
    try {
        const config = msalInstance.getConfiguration();
        if (!config) {
            console.log('MSAL instance exists but not initialized, initializing...');
            await msalInstance.initialize();
        }
    } catch (error) {
        console.warn('Error checking MSAL configuration, re-initializing...', error);
        await msalInstance.initialize();
    }
    return msalInstance;
}

/**
 * Get current account
 */
export async function getCurrentAccount(): Promise<AccountInfo | null> {
    const instance = await getMsalInstance();
    const accounts = instance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
}

/**
 * Get access token for Dataverse
 * First checks localStorage, then MSAL
 */
export async function getAccessToken(): Promise<string | null> {
    try {
        // Get MSAL instance and account first to ensure MSAL is initialized
        const instance = await getMsalInstance();
        const account = await getCurrentAccount();

        if (!account) {
            console.warn('[Auth Service] No account found. User needs to login.');
            // Clear any stored token if account is missing
            clearStoredToken();
            return null;
        }

        // Check localStorage for cached token
        const storedToken = getStoredToken();
        if (storedToken) {
            console.log('[Auth Service] Token retrieved from localStorage');
            // Validate token is not expired and account still exists
            const expiryTime = localStorage.getItem(TOKEN_EXPIRY_KEY);
            if (expiryTime) {
                const expiry = parseInt(expiryTime, 10);
                const timeUntilExpiry = expiry - Date.now();
                // If token expires in less than 5 minutes, refresh it proactively
                if (timeUntilExpiry < 5 * 60 * 1000) {
                    console.log('[Auth Service] Token expires soon, refreshing...');
                    // Fall through to get fresh token
                } else {
                    console.log(`[Auth Service] Token valid, expires in ${Math.floor(timeUntilExpiry / 1000)} seconds`);
                    return storedToken;
                }
            } else {
                return storedToken;
            }
        }

        // If no stored token or expired, get from MSAL
        console.log('[Auth Service] Getting fresh token from MSAL...');
        const request: SilentRequest = {
            ...dataverseRequest,
            account: account,
        };

        try {
            console.log('[Auth Service] Attempting silent token acquisition with scope:', dataverseRequest.scopes);
            const response: AuthenticationResult = await instance.acquireTokenSilent(request);
            console.log('[Auth Service] Token acquired silently from MSAL');
            // Store token in localStorage
            if (response.accessToken && response.expiresOn) {
                const expiresIn = Math.floor((response.expiresOn.getTime() - Date.now()) / 1000);
                storeToken(response.accessToken, expiresIn);
                console.log(`[Auth Service] Token stored in localStorage, expires in ${expiresIn} seconds`);
                console.log('[Auth Service] Token preview:', response.accessToken.substring(0, 20) + '...');
            }
            return response.accessToken;
        } catch (silentError: any) {
            console.log('[Auth Service] Silent token acquisition failed:', silentError.errorCode);
            console.log('[Auth Service] Error message:', silentError.message);
            // If silent token acquisition fails, try interactive popup
            if (silentError.errorCode === 'interaction_required' || 
                silentError.errorCode === 'consent_required' ||
                silentError.errorCode === 'login_required' ||
                silentError.errorCode === 'token_expired') {
                console.log('[Auth Service] Silent token acquisition failed, trying popup...');
                try {
                    const popupResponse = await instance.acquireTokenPopup(dataverseRequest);
                    console.log('[Auth Service] Token acquired via popup from MSAL');
                    // Store token in localStorage
                    if (popupResponse.accessToken && popupResponse.expiresOn) {
                        const expiresIn = Math.floor((popupResponse.expiresOn.getTime() - Date.now()) / 1000);
                        storeToken(popupResponse.accessToken, expiresIn);
                        console.log(`[Auth Service] Token stored in localStorage, expires in ${expiresIn} seconds`);
                        console.log('[Auth Service] Token preview:', popupResponse.accessToken.substring(0, 20) + '...');
                    }
                    return popupResponse.accessToken;
                } catch (popupError: any) {
                    console.error('[Auth Service] Error with popup token acquisition:', popupError);
                    console.error('[Auth Service] Popup error details:', {
                        errorCode: popupError.errorCode,
                        errorMessage: popupError.message,
                        name: popupError.name
                    });
                    return null;
                }
            }
            console.error('[Auth Service] Silent token acquisition error not recoverable:', silentError);
            throw silentError;
        }
    } catch (error: any) {
        console.error('[Auth Service] Error acquiring token:', error);
        console.error('[Auth Service] Error details:', {
            errorCode: error?.errorCode,
            errorMessage: error?.message,
            name: error?.name
        });
        return null;
    }
}

/**
 * Login user
 */
export async function login(): Promise<AuthenticationResult | null> {
    try {
        console.log('Initializing MSAL instance for login...');
        const instance = await getMsalInstance();
        console.log('MSAL instance ready, attempting popup login...');
        console.log('Login request config:', {
            scopes: dataverseRequest.scopes,
            authority: msalConfig.auth.authority
        });
        
        const response = await instance.loginPopup(dataverseRequest);
        console.log('Login popup completed successfully');
        console.log('Login response:', {
            account: response.account ? 'Present' : 'Missing',
            accessToken: response.accessToken ? 'Present' : 'Missing',
            expiresOn: response.expiresOn
        });
        return response;
    } catch (error: any) {
        console.error('Login error:', error);
        console.error('Error details:', {
            errorCode: error.errorCode,
            errorMessage: error.message,
            name: error.name,
            stack: error.stack
        });
        
        // Provide more specific error messages
        if (error.errorCode === 'user_cancelled') {
            console.warn('User cancelled the login popup');
        } else if (error.errorCode === 'popup_window_error') {
            console.error('Popup window error - check if popups are blocked');
        } else if (error.errorCode === 'interaction_in_progress') {
            console.error('Another interaction is already in progress');
        }
        
        return null;
    }
}

/**
 * Logout user
 */
export async function logout(): Promise<void> {
    try {
        // Clear stored token
        clearStoredToken();
        
        const instance = await getMsalInstance();
        const account = await getCurrentAccount();
        if (account) {
            await instance.logoutPopup({ account });
        }
    } catch (error) {
        console.error('Logout error:', error);
        // Clear token even if logout fails
        clearStoredToken();
    }
}

/**
 * Check if user is authenticated
 */
export async function isAuthenticated(): Promise<boolean> {
    const account = await getCurrentAccount();
    if (!account) return false;
    
    const token = await getAccessToken();
    return token !== null;
}
