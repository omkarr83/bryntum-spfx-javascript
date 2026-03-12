/**
 * Dataverse and MSAL configuration for SPFx web part.
 * Update these values to match your Azure AD app and Dataverse environment.
 *
 * Azure AD app registration checklist (Azure Portal > App registrations > Your app):
 * 1. Authentication > Platform configurations: Add "Single-page application".
 * 2. Redirect URIs: Add the exact URL(s) where the web part runs (no query, no hash):
 *    - Workbench: https://mashira365.sharepoint.com/_layouts/15/workbench.aspx
 *    - Site page: https://mashira365.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx
 * 3. Under "Implicit grant and hybrid flows": leave unchecked (SPA uses auth code + PKCE).
 * 4. API permissions: Add Dataverse (e.g. your org's Dynamics CRM) with scope and grant admin consent if required.
 */
export const dataverseConfig = {
  /** Dataverse environment URL (e.g. https://yourorg.crm8.dynamics.com) - no trailing slash */
  // environmentUrl: 'https://orgab553a6a.crm8.dynamics.com',
  environmentUrl: 'https://org77e40fae.crm.dynamics.com',
  /** OData table/entity set for project tasks */
  tableName: 'eppm_projecttasks',
  /** Not used by this web part: import/export run in the browser (Dataverse + MSPDI in mspdiBrowser.ts). Kept only for reference if you add other features that need a backend. */
  apiBaseUrl: ''
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
  // window.location.origin;
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
