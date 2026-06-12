// ============================================================
// connector.js — Conexión MS365 / SharePoint / Graph API
// GESTIONADO POR: Departamento de TI
// NO MODIFICAR salvo cambios de credenciales o fuente de datos
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

msalApp.handleRedirectPromise().then(r => {
  if (r && r.account) { showDashboard(r.account); return; }
  const accs = msalApp.getAllAccounts();
  if (accs.length) showDashboard(accs[0]);
}).catch(e => console.error(e));

function showDashboard(acc) {
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  const n = acc.name || acc.username;
  document.getElementById('uname').textContent = n;
  document.getElementById('uavatar').textContent = n.split(' ').map(x=>x[0]).join('').substring(0,2).toUpperCase();
  loadExcelData();
}

// ── GRAPH API CONFIG ──
const SITE_ID = 'grupoginez.sharepoint.com,eb7ccd99-62a8-4ced-ad00-ac6e27f8b4b3,a1db8bfe-f2dd-4668-a7a5-0189fbad0c61';
const DRIVE_ID = 'b!mc1866hi7UytAKxuJ_i0s_6L26Hd8mhGp6UBifutDGFnXme-tGE-QLyxfALDDF3N';
const SITE_PATH = '/sites/compras.interno';
const FILE_ID = '36D393B2-D8F0-422C-9E91-B4DF3EDBC4F9';
const SHEET_NAME = 'MP';

let RAW = [];

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

    status.textContent = 'Leyendo Excel...';
    const sheetRes = await fetch(
      'https://graph.microsoft.com/v1.0/sites/' + SITE_ID + '/drives/' + DRIVE_ID + '/items/' + FILE_ID + '/workbook/worksheets/' + SHEET_NAME + '/usedRange',
      { headers: { Authorization: 'Bearer ' + token } }
    );
    const sheetData = await sheetRes.json();

    if (!sheetData.values) {
      status.textContent = 'Error leyendo el archivo. Verifica permisos.';
      return;
    }

    status.textContent = 'Procesando datos...';
    RAW = parseExcelData(sheetData.values);
    status.textContent = RAW.length + ' registros cargados desde SharePoint';
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
    var diasAtraso = parseFloat(val(row, 'DIAS DE ATRASO')) || parseFloat(val(row, 'DÍAS DE ATRASO')) || 0;
    var cantPendiente = parseFloat(val(row, 'CANTIDAD PENDIENTE')) || 0;
    var puntoEntrega = String(val(row, 'PUNTO DE ENTREGA') || '').trim();

    if (puntoEntrega === 'Cancelada') continue;
    if (diasAtraso < -100) continue;

    records.push({
      folio: folio,
      fecha: toDateStr(val(row, 'FECHA OC')),
      mes: toDateStr(val(row, 'FECHA OC')).slice(0, 7),
      proveedor: String(val(row, 'PROVEEDOR') || '').trim(),
      producto: String(val(row, 'PRODUCTO') || '').trim(),
      cantidad: parseFloat(val(row, 'CANTIDAD')) || 0,
      um: String(val(row, 'UM') || '').trim(),
      precio_unit: parseFloat(val(row, 'UNITARIO NETO')) || 0,
      total_neto: totalNeto,
      moneda: moneda,
      tc: tc,
      total_mxn: totalMxn,
      condicion: String(val(row, 'CONDICIÓN DE PAGO') || '').trim(),
      dias_atraso: diasAtraso,
      cant_pendiente: cantPendiente,
      cant_entregada: parseFloat(val(row, 'CANTIDAD ENTREGADA')) || 0,
      importe_entregado: parseFloat(val(row, 'IMPORTE DE MAT. ENTREGADO')) || 0,
      total_pagado: parseFloat(val(row, 'PRECIO TOTAL NETO PAGADO')) || 0,
      punto_entrega: puntoEntrega,
      razon_social: String(val(row, 'RAZÓN SOCIAL') || '').trim(),
      folio_factura: String(val(row, 'FOLIO FACTURA O REMISIÓN') || '').trim()
    });
  }
  return records;
}
