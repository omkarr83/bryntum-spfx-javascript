# SPFx development and Chrome workbench fix

## Loading `manifests.js` blocked in Chrome (CORS / Private Network Access)

When you run `gulp serve` and the browser opens the **hosted workbench** at `https://mashira365.sharepoint.com/...`, Chrome can block the request to `https://localhost:4321/temp/manifests.js` with:

- `net::ERR_TIMED_OUT`
- `Access to script at 'http(s)://localhost:4321/temp/manifests.js' from origin 'https://mashira365.sharepoint.com' has been blocked by CORS policy: Permission was denied for this request to access the loopback address space.`

This is due to **Chrome’s Private Network Access (PNA)** rules: a public origin (SharePoint) is not allowed to load scripts from localhost unless the user explicitly allows it.

## Recommended workaround (official approach)

1. Start the dev server **without** opening the workbench:
   ```bash
   npm run serve:debug
   ```
   or:
   ```bash
   gulp serve --nobrowser
   ```

2. Open a **modern SharePoint page** in the same browser and append the debug query string:
   ```
   ?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
   ```
   Example:
   ```
   https://mashira365.sharepoint.com/sites/YourSite/SitePages/Home.aspx?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
   ```
   Use your real site URL (e.g. replace `YourSite` with your site path).

3. When SharePoint shows the **“Load debug scripts”** prompt, click it. The request to localhost is then allowed because it’s triggered by a user action.

4. Add your web part to the page and test. The dev server at `https://localhost:4321` will serve the bundles.

## Other options

- **Edge or Firefox**  
  Use Microsoft Edge or Firefox for SPFx development; they may not enforce PNA the same way, so the hosted workbench might load without this workaround.

- **Chrome: allow local network access**  
  In Chrome (or Edge), for the SharePoint site you can try allowing access to “Local Network” / “Apps on device” so the page can load resources from localhost. Exact steps depend on the browser version.

- **Deploy to app catalog**  
  Build the package (`gulp bundle --ship`), deploy to your tenant app catalog, add the app to a site, and add the web part to a page. Then you are not loading from localhost at all.

---

## Azure AD (MSAL) – avoid redirect loop and “Signing you in…” stuck

The web part uses **popup-first** sign-in when possible so the hosted workbench doesn’t get stuck in a redirect loop. If you still see repeated redirects or “Signing you in…” forever:

1. **Redirect URIs** (Azure Portal → App registration → Authentication → SPA):
   - Add **exactly** the URLs where the web part runs (no query string, no hash), e.g.:
     - `https://mashira365.sharepoint.com/_layouts/15/workbench.aspx`
     - `https://mashira365.sharepoint.com/sites/YourSite/SitePages/Home.aspx` (if you test on a site page)
   - Mismatch causes Microsoft to redirect to a different URL and the token can be lost.

2. **Allow popups** for the SharePoint site so the sign-in popup can open; if popups are blocked, the code falls back to full-page redirect (which can loop if the hash is stripped).

3. **API permissions**: Ensure the app has the right Dataverse/Dynamics CRM permission and admin consent if required.

---

## Import/Export XML – "Failed to fetch"

Import and Export use a **backend API** (the `server` folder). If you see **"Import failed: Failed to fetch"** or **"Export failed: Failed to fetch"**:

1. **Start the backend server**  
   In the `server` folder run:
   ```bash
   npm install
   npm run dev
   ```
   The server runs at `http://localhost:3001` by default.

2. **HTTPS (mixed content)**  
   When the web part runs on **HTTPS** (e.g. `https://mashira365.sharepoint.com`), the browser blocks requests to **HTTP** URLs (e.g. `http://localhost:3001`). So:
   - Either deploy the backend to an **HTTPS** URL (e.g. Azure App Service, or a tunnel like ngrok: `ngrok http 3001`), then set that URL in the web part property **"API base URL"** (or in `dataverseConfig.apiBaseUrl`).
   - Or test the web part on an HTTP page (e.g. local workbench) so both page and API are HTTP.

3. **Web part property**  
   In the workbench or page, edit the web part → **API base URL** and set your backend URL (e.g. `https://your-api.azurewebsites.net/api` or `https://xxxx.ngrok.io/api`). No trailing slash.
