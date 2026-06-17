// ============================================================
// connector.js -- Conexion MS365 / SharePoint / Graph API
// Módulo de autenticación y conexión a fuentes de datos
// ============================================================

const CLIENT_ID = "5cbee9d4-0574-49fb-8d63-fcac84bfa44d";
const TENANT_ID = "d08c56ca-3b55-42db-b365-359cf1503e4e";

const msalCfg = {
  auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/" + TENANT_ID, redirectUri: "https://grupoginez.github.io/dashboard-comprasgg/" },
  cache: { cacheLocation: "localStorage" }
};
const msalApp = new msal.PublicClientApplication(msalCfg);

async function login() {
  try {
    await msalApp.loginRedirect({ scopes: ["User.Read"] });
  } catch(e) {
    document.getElementById('login-err').style.display = 'block';
  }
}
function logout() { msalApp.logoutRedirect(); }

function handleAuthResult(r) {
  const loginEl = document.getElementById('login-section');
  if (!loginEl) {
    setTimeout(() => handleAuthResult(r), 50);
    return;
  }
  if (r && r.account) { showDashboard(r.account); return; }
  const accs = msalApp.getAllAccounts();
  if (accs.length) showDashboard(accs[0]);
}

msalApp.handleRedirectPromise().then(r => {
  handleAuthResult(r);
}).catch(e => console.error(e));
