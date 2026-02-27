# Fix AADSTS9002326: Cross-origin token redemption / Single-Page Application

The error **AADSTS9002326** means Azure AD is rejecting the token request because the app is **not** registered as a **Single-Page Application (SPA)**. It must be fixed in the **Azure Portal**, not in code.

## Steps in Azure Portal

1. Open **Microsoft Entra ID (Azure Active Directory)** → **App registrations** → select your app (the one with client ID used in `dataverseConfig` / `msalConfig`).

2. Go to **Authentication**.

3. Under **Platform configurations**:
   - If you see **Web** and no **Single-page application**:  
     - Click **Add a platform** → choose **Single-page application**.
   - Under **Single-page application**, add **Redirect URIs** that **exactly** match where the app runs:
     - `https://jssca.sharepoint.com`
     - `https://jssca.sharepoint.com/`
     - If you use the workbench: `https://jssca.sharepoint.com/_layouts/15/workbench.aspx`
     - Add any other SharePoint page URLs where you embed the web part (e.g. site pages).

4. **Important**: For this SPA (SharePoint-hosted), do **not** rely only on **Web** platform. The **Single-page application** platform is required for the redirect flow and cross-origin token redemption from `https://jssca.sharepoint.com`.

5. Save the changes. After a short delay, sign-in and token redemption from your SharePoint origin should work.

## Redirect URI in code

The web part uses the current page URL (without query/hash) as `redirectUri`. Ensure that URL is one of the redirect URIs you added in Azure for the **Single-page application** platform. If you prefer a fixed list, you can set in `dataverseConfig.ts` a fixed `redirectUri` (e.g. `https://jssca.sharepoint.com`) and add that exact value in Azure.

## Summary

- **Cause**: App registered as **Web** (or SPA not configured) and/or redirect URI mismatch.
- **Fix**: Add **Single-page application** platform and the correct **Redirect URIs** for your SharePoint site in the app registration.
