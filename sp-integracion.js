/* ============================================================
   sp-integracion.js  ·  Parte Diario Telice
   Capa de LECTURA de SharePoint -> objeto OBRAS de la app
   ------------------------------------------------------------
   Descubrimiento clave (sesión previa):
   SharePoint devuelve los items con nombres internos POSICIONALES:
     · 1ª columna de la lista  -> "Title"
     · resto, por orden        -> "field_1", "field_2", ...
   Los números llegan como número y el texto como texto.

   Por tanto el mapeo es 100% por POSICIÓN, no por nombre de columna.
   SP_SCHEMA define, para cada maestro y EN ORDEN DE COLUMNA:
     índice 0 -> Title ; índice k -> field_k
   y a qué campo de la app va cada uno. '_skip' descarta, '_proyecto'
   se usa para repartir las filas entre obras.
   ============================================================ */
(function (global) {
'use strict';

/* ---------- CONFIG + MSAL (patrón idéntico al admin-listas.html) ---------- */
var _cfg = null, _msal = null;

/* redirectUri: debe ser una URL YA REGISTRADA en Entra, del mismo origen y
   que cargue (200). Usamos la carpeta del repo (la sirve GitHub Pages como
   index.html, que es el redirect registrado), NO la página actual.
   Si el redirect registrado fuese exactamente ".../index.html", cambia
   REDIRECT_URI por esa cadena. */
function carpetaActual() {
  var base = window.location.pathname.replace(/\/[^/]*$/, '/'); // quita el fichero
  return window.location.origin + base;                         // .../parte-diario-telice/
}
var REDIRECT_URI = carpetaActual();

async function loadConfig() {
  var base = window.location.pathname.replace(/\/[^/]*$/, '/');
  var url  = window.location.origin + base + 'config.json';
  var r = await fetch(url);
  if (!r.ok) throw new Error('config.json no encontrado en ' + url);
  var c = await r.json();
  if (!c.clientId || !c.tenantId || !c.siteUrl) throw new Error('config.json incompleto');
  return c;
}

function initMsal(cfg) {
  _msal = new msal.PublicClientApplication({
    auth: { clientId: cfg.clientId,
            authority: 'https://login.microsoftonline.com/' + cfg.tenantId,
            redirectUri: REDIRECT_URI },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
  });
}

// Lectura: bastaría AllSites.Read, pero mantenemos Manage para reutilizar
// el consentimiento ya concedido y no disparar una segunda pantalla de login.
function getScopes() { return [new URL(_cfg.siteUrl).origin + '/AllSites.Manage']; }

async function ensureReady() {
  if (!_cfg)  _cfg  = await loadConfig();
  if (!_msal) initMsal(_cfg);
}

async function login() {
  await ensureReady();
  var r = await _msal.loginPopup({ scopes: getScopes() });
  return r.account;
}

function currentAccount() {
  var ac = _msal.getAllAccounts();
  return ac.length ? ac[0] : null;
}

async function getToken() {
  var ac = _msal.getAllAccounts();
  if (!ac.length) throw new Error('Sin sesión activa');
  try { return (await _msal.acquireTokenSilent({ scopes: getScopes(), account: ac[0] })).accessToken; }
  catch (e) { return (await _msal.acquireTokenPopup({ scopes: getScopes() })).accessToken; }
}

/* ---------- REST helper (idéntico al admin) ---------- */
function spUrl(path) { return _cfg.siteUrl.replace(/\/$/, '') + '/_api/web' + path; }

async function spGet(path) {
  var t = await getToken();
  var r = await fetch(spUrl(path), {
    headers: { 'Authorization': 'Bearer ' + t, 'Accept': 'application/json;odata=nometadata' }
  });
  if (!r.ok) {
    var msg = 'HTTP ' + r.status;
    try { var j = await r.json(); if (j.error && j.error.message) msg += ': ' + j.error.message.value; } catch (e) {}
    throw new Error(msg);
  }
  return r.json();
}

// Lee TODOS los items de una lista ($top 5000 sobra para los maestros).
async function spReadAll(listTitle) {
  var out = [];
  var path = "/lists/GetByTitle('" + listTitle + "')/items?$top=5000";
  var data = await spGet(path);
  if (data && data.value) out = out.concat(data.value);
  return out;
}

/* ---------- ESQUEMA POSICIONAL ----------
   cols[i] = campo destino en la app para la columna i de la lista
   i === 0  -> 'Title'      ;  i >= 1 -> 'field_i'
   '_skip'      -> se ignora
   '_proyecto'  -> código de obra (reparte la fila)
   Para 'proyecto' (M_Proyectos) los destinos son campos de cabecera de obra. */
var SP_SCHEMA = {
  M_Proyectos: {
    kind: 'proyecto',
    cols: ['codigo','nombre','tipo','incentivos']
    //      Title    f1      f2     f3
  },
  M_Partidas: {
    kind: 'array', target: 'partidas',
    cols: ['_skip','_proyecto','cod','desc','ud','pu','coste_ud','te','ts','ti','to','_grupo','_tipo_grupo']
    //      Title   f1          f2    f3     f4   f5   f6         f7   f8   f9   f10  f11      f12
  },
  M_PartidasCert: {
    kind: 'array', target: 'partidas_cert',
    cols: ['_skip','_proyecto','cod','desc','ud','med','pu','venta_ppto','coste_ud','cap','te','ts','ti','to']
  },
  M_PartidasPpto: {
    kind: 'array', target: 'partidas_ppto',
    cols: ['_skip','_proyecto','cod','desc','ud','med','pu_ec','coste_ud','venta_ppto','coste_total_ppto','te','ts','ti','to']
  },
  M_Personal: {
    kind: 'array', target: 'personal',
    cols: ['_skip','_proyecto','num','nombre','cat','tipo','ch','dc','md','km']
  },
  M_Materiales: {
    kind: 'array', target: 'mat',
    cols: ['_skip','_proyecto','cod','desc','ud','pu','coste_ud','te','ts','_grupo'],
    defaults: { ti: '', to: '', _tipo_grupo: 'SUM' }  // la app espera estos; el maestro no los trae
  },
  M_Subcontratas: {
    kind: 'array', target: 'sub',
    cols: ['_skip','_proyecto','task','cod','desc','ud','coste_estudio','tarifa','empresa','contrato']
  }
};

// Campos que la app trata como número (se fuerza Number por seguridad).
var NUM_FIELDS = {
  pu:1, coste_ud:1, med:1, venta_ppto:1, pu_ec:1, coste_total_ppto:1,
  ch:1, dc:1, md:1, km:1, coste_estudio:1, tarifa:1
};

function internalName(i) { return i === 0 ? 'Title' : 'field_' + i; }

function num(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

// Mapea una fila cruda de SP -> objeto con los campos de la app.
function mapRow(raw, schema) {
  var o = {};
  var cols = schema.cols;
  for (var i = 0; i < cols.length; i++) {
    var dest = cols[i];
    if (dest === '_skip') continue;
    var val = raw[internalName(i)];
    if (val === undefined) val = '';
    if (dest === '_proyecto') { o._proyecto = (val == null ? '' : String(val)).trim(); continue; }
    o[dest] = NUM_FIELDS[dest] ? num(val) : (val == null ? '' : val);
  }
  if (schema.defaults) for (var k in schema.defaults) if (o[k] === undefined) o[k] = schema.defaults[k];
  return o;
}

/* ---------- CONSTRUCCIÓN DE OBRAS ---------- */
function nuevaObra(cab) {
  return {
    codigo: cab.codigo, nombre: cab.nombre, tipo: cab.tipo, incentivos: cab.incentivos,
    partidas: [], partidas_cert: [], partidas_ppto: [],
    personal: [], mat: [], sub: [], maq_obra: []
  };
}

// Lee todos los maestros y construye el objeto OBRAS completo.
async function construirObras(onProgress) {
  await ensureReady();
  var log = onProgress || function(){};
  var OBRAS = {};

  // 1) Cabeceras de obra desde M_Proyectos
  log('Leyendo M_Proyectos…');
  var proy = await spReadAll('M_Proyectos');
  proy.forEach(function (raw) {
    var cab = mapRow(raw, SP_SCHEMA.M_Proyectos);
    if (cab.codigo) OBRAS[String(cab.codigo).trim()] = nuevaObra({
      codigo: String(cab.codigo).trim(), nombre: cab.nombre || '',
      tipo: cab.tipo || '', incentivos: cab.incentivos || 'ninguno'
    });
  });
  log('Obras encontradas: ' + Object.keys(OBRAS).join(', '));

  // 2) Resto de maestros (arrays), repartidos por _proyecto
  var arrays = ['M_Partidas','M_PartidasCert','M_PartidasPpto','M_Personal','M_Materiales','M_Subcontratas'];
  for (var a = 0; a < arrays.length; a++) {
    var name = arrays[a];
    var sch  = SP_SCHEMA[name];
    log('Leyendo ' + name + '…');
    var rows = await spReadAll(name);
    var huerfanas = 0;
    rows.forEach(function (raw) {
      var item = mapRow(raw, sch);
      var cod  = item._proyecto;
      delete item._proyecto;
      var obra = OBRAS[cod];
      if (!obra) { huerfanas++; return; }
      obra[sch.target].push(item);
    });
    log('  ' + name + ': ' + rows.length + ' filas' + (huerfanas ? ' (' + huerfanas + ' sin obra, ignoradas)' : ''));
  }
  return OBRAS;
}

/* ---------- APLICAR a la app ----------
   OBRAS en index.html es 'const': mutamos el objeto in situ. */
function aplicarObras(destino, nuevas) {
  Object.keys(destino).forEach(function (k) { delete destino[k]; });
  Object.keys(nuevas).forEach(function (k) { destino[k] = nuevas[k]; });
  return destino;
}

/* ---------- API pública ---------- */
global.SPTelice = {
  ensureReady: ensureReady,
  login: login,
  currentAccount: currentAccount,
  construirObras: construirObras,
  aplicarObras: aplicarObras,
  spGet: spGet,
  SP_SCHEMA: SP_SCHEMA
};

})(window);
