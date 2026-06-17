// ============================================================
// connector.js -- Conexion MS365 / SharePoint / Graph API
// Módulo de autenticación y conexión a fuentes de datos
// ============================================================

const CLIENT_ID = "5cbee9d4-0574-49fb-8d63-fcac84bfa44d";
const TENANT_ID = "d08c56ca-3b55-42db-b365-359cf1503e4e";

const msalCfg = {
  auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/" + TENANT_ID, redirectUri: "https://grupoginez.github.io/dashboard-comprasgg/" },
  cache: { cacheLocation: "sessionStorage" }
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
    // DOM aun no listo, reintentar
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

function showDashboard(acc) {
  document.getElementById('login-section').style.display = 'none';
  document.getElementById('dashboard-section').style.display = 'block';
  const n = acc.name || acc.username;
  document.getElementById('uname').textContent = n;
  document.getElementById('uavatar').textContent = n.split(' ').map(x=>x[0]).join('').substring(0,2).toUpperCase();
  loadExcelData();
}

// -- GRAPH API CONFIG --
const SITE_ID = 'grupoginez.sharepoint.com,eb7ccd99-62a8-4ced-ad00-ac6e27f8b4b3,a1db8bfe-f2dd-4668-a7a5-0189fbad0c61';
const DRIVE_ID = 'b!mc1866hi7UytAKxuJ_i0s_6L26Hd8mhGp6UBifutDGFnXme-tGE-QLyxfALDDF3N';
const FILE_ID = 'F6B14DAE-88ED-4CAD-B9A5-D9D3BCBF9509';
const SHEET_NAMES = ['MP', 'PIGMENTOS', 'FRAGANCIAS', 'PIPAS', 'OTROS'];

// RAW_BY_SHEET = { MP: [...], PIGMENTOS: [...], FRAGANCIAS: [...], PIPAS: [...], OTROS: [...] }
let RAW_BY_SHEET = {};

async function getGraphToken() {
  const req = { scopes: ['https://graph.microsoft.com/Files.Read.All'], account: msalApp.getAllAccounts()[0] };
  try {
    const r = await msalApp.acquireTokenSilent(req);
    return r.accessToken;
  } catch(e) {
    await msalApp.acquireTokenRedirect(req);
    return null;
  }
}

async function loadExcelData() {
  const status = document.getElementById('load-status');
  try {
    status.textContent = 'Conectando con SharePoint...';
    const token = await getGraphToken();
    if (!token) return;

    let totalRegistros = 0;

    for (const sheet of SHEET_NAMES) {
      status.textContent = 'Leyendo hoja ' + sheet + '...';
      const sheetRes = await fetch(
        'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/drives/' + DRIVE_ID + '/items/' + FILE_ID + '/workbook/worksheets/' + sheet + '/usedRange',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      const sheetData = await sheetRes.json();

      if (!sheetData.values) {
        console.error('Error leyendo hoja', sheet, sheetData.error);
        RAW_BY_SHEET[sheet] = [];
        continue;
      }

      RAW_BY_SHEET[sheet] = parseExcelData(sheetData.values);
      totalRegistros += RAW_BY_SHEET[sheet].length;
    }

    status.textContent = totalRegistros + ' registros cargados desde SharePoint';
    setTimeout(function() { status.textContent = ''; }, 3000);
    initDashboard();

  } catch(err) {
    status.textContent = 'Error: ' + err.message;
    console.error(err);
  }
}

function toDateStr(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    var d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  return String(val).slice(0, 10);
}

// ============================================================
// PARSER: traduce filas del Excel a objetos JS
// Estructura comun a las 5 hojas: headers en fila indice 3,
// datos desde fila indice 4. PIPAS no tiene PRESENTACION.
// Si cambian columnas del Excel, ajustar este parser
// ============================================================
function parseExcelData(values) {
  var headers = values[3].map(function(h) { return String(h).trim(); });
  var colIdx = {};
  headers.forEach(function(h, i) { colIdx[h] = i; });

  function col(name) { return colIdx[name] !== undefined ? colIdx[name] : -1; }
  function val(row, name) { var i = col(name); return i >= 0 ? row[i] : ''; }

  var records = [];
  for (var i = 4; i < values.length; i++) {
    var row = values[i];
    var folio = String(val(row, 'FOLIO OC') || '').trim();
    if (!folio) continue;

    var moneda = String(val(row, 'MONEDA') || 'MXN').trim();
    var totalNeto = parseFloat(val(row, 'TOTAL NETO')) || 0;
    var tc = parseFloat(val(row, 'TIPO DE CAMBIO (SOLICITUD)')) || (moneda === 'USD' ? 18.5 : 1);
    var totalMxn = moneda === 'USD' ? totalNeto * tc : totalNeto;
    var diasAtraso = parseFloat(val(row, 'DIAS DE ATRASO')) || parseFloat(val(row, 'D\u00cdAS DE ATRASO')) || 0;
    var cantPendiente = parseFloat(val(row, 'CANTIDAD PENDIENTE')) || 0;
    var puntoEntrega = String(val(row, 'PUNTO DE ENTREGA') || '').trim();

    if (puntoEntrega === 'Cancelada') continue;
    if (diasAtraso < -100) continue;

    records.push({
      no_requi: String(val(row, 'NO. REQUI') || '').trim(),
      folio: folio,
      fecha: toDateStr(val(row, 'FECHA OC')),
      mes: toDateStr(val(row, 'FECHA OC')).slice(0, 7),
      proveedor: String(val(row, 'PROVEEDOR') || '').trim(),
      producto: String(val(row, 'PRODUCTO') || '').trim(),
      cantidad: parseFloat(val(row, 'CANTIDAD')) || 0,
      um: String(val(row, 'UM') || '').trim(),
      precio_unit: parseFloat(val(row, 'PRECIO UNITARIO')) || 0,
      iva: parseFloat(val(row, 'IVA')) || 0,
      unitario_neto: parseFloat(val(row, 'UNITARIO NETO')) || 0,
      total_neto: totalNeto,
      moneda: moneda,
      tc: tc,
      total_mxn: totalMxn,
      importe_entregado: parseFloat(val(row, 'IMPORTE DE MAT. ENTREGADO')) || 0,
      presentacion: String(val(row, 'PRESENTACI\u00d3N') || '').trim(),
      fecha_sol_entrega: toDateStr(val(row, 'FECHA SOLICITADA DE ENTREGA')) || toDateStr(val(row, 'FECHA SOL. DE ENTREGA')),
      fecha_entrega_real: toDateStr(val(row, 'FECHA ENTREGA REAL')),
      dias_atraso: diasAtraso,
      cant_entregada: parseFloat(val(row, 'CANTIDAD ENTREGADA')) || 0,
      cant_pendiente: cantPendiente,
      punto_entrega: puntoEntrega,
      razon_social: String(val(row, 'RAZ\u00d3N SOCIAL') || '').trim(),
      folio_factura: String(val(row, 'FOLIO FACTURA O REMISI\u00d3N') || '').trim(),
      estatus: String(val(row, 'ESTATUS') || '').trim(),
      condicion: String(val(row, 'CONDICI\u00d3N DE PAGO') || '').trim()
    });
  }
  return records;
}
