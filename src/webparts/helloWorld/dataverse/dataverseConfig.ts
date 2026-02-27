/**
 * Dataverse and MSAL configuration for SPFx web part.
 * Update these values to match your Azure AD app and Dataverse environment.
 */
export const dataverseConfig = {
  /** Dataverse environment URL (e.g. https://yourorg.crm8.dynamics.com) - no trailing slash */
  // environmentUrl: 'https://orgab553a6a.crm8.dynamics.com',
  environmentUrl: 'https://org77e40fae.crm.dynamics.com',
  /** OData table/entity set for project tasks */
  tableName: 'eppm_projecttasks',
  /** Backend API base URL for Export/Import MS Project (e.g. http://localhost:3001/api) - no trailing slash */
  apiBaseUrl: 'http://localhost:3001/api'
};

/**
 * MSAL configuration for SPFx (SharePoint origin).
 * Required for AADSTS9002326 fix: In Azure Portal, register this app as
 * "Single-page application" and add redirectUri (e.g. https://jssca.sharepoint.com)
 * under Authentication > Platform configurations > SPA.
 */
function getRedirectUri(): string {
  if (typeof window === 'undefined') return '';
  // Use full URL so redirect returns to the same page (workbench or site page)
  return window.location.href.split('?')[0].split('#')[0];
}

export const msalConfig = {
  auth: {
    // clientId: 'd7cedaf0-7f7e-4779-9985-37d8ac9fb8c0',
    // authority: 'https://login.microsoftonline.com/cf50b276-a7b3-4cd0-bd1f-a3a13316b1a5',
    clientId: '88994ec9-c137-4cb2-828a-2e0ed12ccaf0',
    authority: 'https://login.microsoftonline.com/6cc11140-3317-48cd-99d9-25abe8e51d67',
    redirectUri: getRedirectUri()
  },
  cache: {
    cacheLocation: 'localStorage' as const,
    storeAuthStateInCookie: false
  }
};

/** Scope for Dataverse API access (resource/.default) */
export const dataverseScope = `${dataverseConfig.environmentUrl}/.default`;
