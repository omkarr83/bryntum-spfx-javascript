import { Configuration, PopupRequest } from '@azure/msal-browser';

// MSAL configuration
// IMPORTANT: The redirectUri must match exactly what's configured in Azure AD
// For local development, use: http://localhost:5173 (or the port Vite assigns)
// For production, update this to your production URL
// 
// NOTE: When using popup flow, Azure AD still validates the redirect URI
// Make sure to add BOTH with and without trailing slash in Azure AD:
// - http://localhost:5173
// - http://localhost:5173/

// Get the current origin (e.g., http://localhost:5173)
const getRedirectUri = (): string => {
    // Use the current origin (works for both dev and production)
    const origin = window.location.origin;
    console.log('MSAL redirect URI:', origin);
    return origin;
};

export const msalConfig: Configuration = {
    auth: {
        clientId: 'd7cedaf0-7f7e-4779-9985-37d8ac9fb8c0',
        authority: 'https://login.microsoftonline.com/cf50b276-a7b3-4cd0-bd1f-a3a13316b1a5',
        // Use dynamic redirect URI based on environment
        redirectUri: getRedirectUri(),
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false,
    },
};

// Add scopes here for ID token to be used at Microsoft identity platform endpoints.
export const loginRequest: PopupRequest = {
    scopes: ['User.Read'],
};

// Add the client app ID as a scope for the Dataverse resource
export const dataverseRequest: PopupRequest = {
    scopes: ['https://orgab553a6a.crm8.dynamics.com/.default'],
};

export const dataverseResource = 'https://orgab553a6a.crm8.dynamics.com/';
