/**
 * graph.js
 * Handles Microsoft sign-in (MSAL.js) and talking to Microsoft Graph so the
 * app can create/update the Excel workbook (and optionally photos) inside a
 * dedicated OneDrive folder.
 *
 * SETUP REQUIRED BEFORE THIS WILL WORK — see README.md "OneDrive / Excel
 * configuration" section. You must register an app in Azure AD (Entra ID)
 * and fill in CLIENT_ID below. Nothing here will function with the
 * placeholder values.
 */

const GRAPH_CONFIG = {
  // ---- FILL THESE IN (see README.md) ----
  clientId: 'YOUR-AZURE-AD-APP-CLIENT-ID',
  authority: 'https://login.microsoftonline.com/common', // use your tenant ID instead of "common" to restrict to your org
  redirectUri: window.location.origin + window.location.pathname,
  // -----------------------------------------

  scopes: ['Files.ReadWrite', 'User.Read', 'offline_access'],
  folderName: 'Verizon North Royalton UPS - Americano Checklists',
  workbookName: 'Verizon_NorthRoyalton_UPS_Inspections.xlsx',
};

let msalInstance = null;
let activeAccount = null;

function isConfigured() {
  return GRAPH_CONFIG.clientId && !GRAPH_CONFIG.clientId.startsWith('YOUR-');
}

function getMsal() {
  if (!msalInstance) {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: GRAPH_CONFIG.clientId,
        authority: GRAPH_CONFIG.authority,
        redirectUri: GRAPH_CONFIG.redirectUri,
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: true, // helps on iOS Safari (ITP can drop localStorage-only auth state)
      },
    });
  }
  return msalInstance;
}

const GraphSync = {
  isConfigured,

  isSignedIn() {
    return !!activeAccount;
  },

  getAccountName() {
    return activeAccount ? (activeAccount.name || activeAccount.username) : null;
  },

  async init() {
    if (!isConfigured()) return false;
    const client = getMsal();
    await client.handleRedirectPromise().catch(() => null);
    const accounts = client.getAllAccounts();
    if (accounts.length > 0) {
      activeAccount = accounts[0];
      return true;
    }
    return false;
  },

  async signIn() {
    if (!isConfigured()) {
      throw new Error('Microsoft Graph is not configured yet. Add your Azure AD Client ID in js/graph.js.');
    }
    const client = getMsal();
    // Popup is generally more reliable than redirect inside iOS home-screen PWAs.
    const result = await client.loginPopup({ scopes: GRAPH_CONFIG.scopes });
    activeAccount = result.account;
    return activeAccount;
  },

  async signOut() {
    if (!msalInstance || !activeAccount) return;
    await msalInstance.logoutPopup({ account: activeAccount });
    activeAccount = null;
  },

  async getToken() {
    const client = getMsal();
    const request = { scopes: GRAPH_CONFIG.scopes, account: activeAccount };
    try {
      const res = await client.acquireTokenSilent(request);
      return res.accessToken;
    } catch (e) {
      const res = await client.acquireTokenPopup(request);
      activeAccount = res.account;
      return res.accessToken;
    }
  },

  async graphFetch(path, options = {}) {
    const token = await this.getToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(options.headers || {}),
      },
    });
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      const err = new Error(`Graph API error ${res.status}: ${text}`);
      err.status = res.status;
      throw err;
    }
    return res;
  },

  /** Ensures the dedicated OneDrive folder exists (creates it on first run). */
  async ensureFolder() {
    const safeName = encodeURIComponent(GRAPH_CONFIG.folderName);
    try {
      await this.graphFetch(`/me/drive/root:/${safeName}`);
      return true;
    } catch (e) {
      if (e.status && e.status !== 404) throw e; // auth/permission errors should surface, not be swallowed
      // Not found -> create it
      await this.graphFetch('/me/drive/root/children', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: GRAPH_CONFIG.folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename',
        }),
      });
      return true;
    }
  },

  /** Uploads (creates or overwrites) the workbook file as raw bytes. Simple PUT — fine up to ~4MB. */
  async uploadWorkbook(arrayBuffer) {
    await this.ensureFolder();
    const path = `${GRAPH_CONFIG.folderName}/${GRAPH_CONFIG.workbookName}`;
    const safePath = path.split('/').map(encodeURIComponent).join('/');
    const res = await this.graphFetch(`/me/drive/root:/${safePath}:/content`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: arrayBuffer,
    });
    return res.json();
  },

  /** Uploads a single inspection photo into a Photos subfolder and returns its shareable webUrl. */
  async uploadPhoto(blob, filename) {
    await this.ensureFolder();
    const path = `${GRAPH_CONFIG.folderName}/Photos/${filename}`;
    const safePath = path.split('/').map(encodeURIComponent).join('/');
    const res = await this.graphFetch(`/me/drive/root:/${safePath}:/content`, {
      method: 'PUT',
      headers: { 'Content-Type': blob.type || 'image/jpeg' },
      body: blob,
    });
    const json = await res.json();
    return json.webUrl;
  },

  config: GRAPH_CONFIG,
};

window.GraphSync = GraphSync;
