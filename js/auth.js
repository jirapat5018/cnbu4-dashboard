/**
 * auth.js - Microsoft Entra ID sign-in with MSAL.js (Authorization Code + PKCE, public client, NO secret).
 * Tokens live in sessionStorage managed by MSAL; nothing is hard-coded.
 */
export class Auth {
  constructor(cfg) {
    this.cfg = cfg;
    this.msal = new window.msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
        redirectUri: cfg.redirectUri,
        postLogoutRedirectUri: cfg.redirectUri,
      },
      cache: { cacheLocation: 'sessionStorage' },
    });
    this.account = null;
  }

  async init() {
    await this.msal.initialize();
    const result = await this.msal.handleRedirectPromise();
    this.account = result?.account || this.msal.getAllAccounts()[0] || null;
    if (this.account) this.msal.setActiveAccount(this.account);
    return this.account;
  }

  signIn() { return this.msal.loginRedirect({ scopes: this.cfg.scopes }); }
  signOut() { return this.msal.logoutRedirect({ account: this.account }); }
  get userName() { return this.account?.name || this.account?.username || ''; }

  async getToken() {
    if (!this.account) throw new Error('ยังไม่ได้ Sign in');
    try {
      return (await this.msal.acquireTokenSilent({ scopes: this.cfg.scopes, account: this.account })).accessToken;
    } catch (e) {
      if (e instanceof window.msal.InteractionRequiredAuthError) { await this.msal.acquireTokenRedirect({ scopes: this.cfg.scopes }); }
      throw e;
    }
  }
}
