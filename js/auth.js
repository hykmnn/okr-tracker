// Đăng nhập Microsoft (MSAL redirect flow — giống TA dashboard; popup hay bị
// trình duyệt chặn). Token SharePoint delegated: người dùng chỉ đọc/ghi được
// những gì chính họ có quyền trên site.
import { CONFIG } from "./config.js";

const MSAL_CDN = "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.30.0/lib/msal-browser.min.js";
export const SCOPES = [`https://${CONFIG.siteHost}/AllSites.Write`];
let app = null;

function loadMsal() {
  return new Promise((resolve, reject) => {
    if (window.msal) return resolve();
    const s = document.createElement("script");
    s.src = MSAL_CDN;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Không tải được thư viện đăng nhập Microsoft"));
    document.head.appendChild(s);
  });
}

export async function initAuth() {
  await loadMsal();
  app = new window.msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: location.origin + location.pathname,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await app.initialize();
  const r = await app.handleRedirectPromise();
  return r?.account || app.getAllAccounts()[0] || null;
}

export async function signIn() {
  if (!app) await initAuth();
  await app.loginRedirect({ scopes: SCOPES });
}

export async function signOut() {
  const account = app?.getAllAccounts()[0];
  await app.logoutRedirect({ account, postLogoutRedirectUri: location.origin + location.pathname });
}

export async function getToken() {
  const account = app?.getAllAccounts()[0];
  if (!account) throw new Error("Chưa đăng nhập");
  try {
    return (await app.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
  } catch {
    await app.acquireTokenRedirect({ scopes: SCOPES, account });
    return new Promise(() => {}); // trang sẽ chuyển hướng
  }
}

export function currentAccount() {
  const a = app?.getAllAccounts()[0];
  return a ? { name: a.name || a.username, username: (a.username || "").toLowerCase() } : null;
}
