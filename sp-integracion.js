/* ============================================================
   sp-integracion.js  ·  Parte Diario Telice
   Capa de LECTURA de SharePoint  (modelo fiel de 11 maestros)
   ------------------------------------------------------------
   Regla posicional confirmada con los Excel origen:
     1ª columna -> "Title" ; resto, por orden -> field_1, field_2...
   SP_SCHEMA.cols[i] = nombre real de la columna i de cada lista
     (i===0 -> Title ; i>=1 -> field_i)
   'num'  = columnas a forzar a número.
   Salida: objeto DATOS con un array por maestro, con los nombres
   reales de columna. NO se fuerza a la estructura OBRAS de la app
   (eso se reconcilia en un paso posterior).
   ============================================================ */
(function (global) {
'use strict';

/* ---------- CONFIG + MSAL ---------- */
var _cfg = null, _msal = null;

/* redirectUri = carpeta del repo (servida como index.html, ya registrada
   en Entra). Mismo origen + 200, suficiente para el relay del popup. */
function carpetaActual() {
  var base = window.location.pathname.replace(/\/[^/]*$/, '/');
  return window.location.origin + base;
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

/* ---------- REST helper ---------- */
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

async function spReadAll(listTitle) {
  var data = await spGet("/lists/GetByTitle('" + listTitle + "')/items?$top=5000");
  return (data && data.value) ? data.value : [];
}

// Devuelve un mapa { TituloColumna: InternalNameReal } leyendo los campos de la lista.
// Esto desacopla la lectura del ORDEN de creación de columnas (field_N): leemos por
// nombre, no por posición, así reordenar o añadir columnas no rompe nada.
async function spFieldMap(listTitle) {
  var data = await spGet("/lists/GetByTitle('" + listTitle +
    "')/fields?$select=Title,InternalName,Hidden,ReadOnlyField&$filter=Hidden eq false");
  var map = {};
  (data.value || []).forEach(function (f) {
    // Title es la primera columna real (InternalName === 'Title'); el resto por su Title.
    if (map[f.Title] === undefined) map[f.Title] = f.InternalName;
  });
  return map;
}

/* ---------- ESQUEMA POSICIONAL (11 maestros, nombres reales) ---------- */
var SP_SCHEMA = {
  // --- Corporativos (sin columna Proyecto) ---
  M_Proyectos: { target:'proyectos', proyecto:false,
    cols:['Clave','Codigo','Nombre','Tipo','Cliente','Estado','Incentivos'], num:[] },
  M_Personal: { target:'personal', proyecto:false,
    cols:['Clave','Num','Nombre','Categoria','Tipo','CosteHora','DietaCompleta','MediaDieta'],
    num:['Num','CosteHora','DietaCompleta','MediaDieta'] },
  M_Maquinaria: { target:'maquinaria', proyecto:false,
    cols:['Clave','Codigo','Descripcion','Categoria','CosteHora','CosteDia','Observaciones'],
    num:['CosteHora','CosteDia'] },
  // --- Por obra (Proyecto en field_1) ---
  M_Partidas: { target:'partidas', proyecto:true,
    cols:['Clave','Proyecto','Codigo','CtrlIngenieria','CtrlSuministro','CtrlEjecucion','CtrlOtros','Capitulo','Ud','Descripcion','Medicion','PrecioVenta','CosteUd','PuVentaIng','CosteUdIng','PuVentaSum','CosteUdSum','PuVentaEje','CosteUdEje','PuVentaOtr','CosteUdOtr'],
    num:['Medicion','PrecioVenta','CosteUd','PuVentaIng','CosteUdIng','PuVentaSum','CosteUdSum','PuVentaEje','CosteUdEje','PuVentaOtr','CosteUdOtr'] },
  M_PartidasControl: { target:'partidasControl', proyecto:true,
    cols:['Clave','Proyecto','Codigo','Capitulo','Descripcion','Ud','Medicion','CosteUnitario','CosteTotal'],
    num:['Medicion','CosteUnitario','CosteTotal'] },
  M_Subcontratas: { target:'subcontratas', proyecto:true,
    cols:['Clave','Proyecto','TaskBC','CodigoPartida','Descripcion','Ud','CosteEstudio','Tarifa','Empresa','Contrato'],
    num:['CosteEstudio','Tarifa'] },
  M_Materiales: { target:'materiales', proyecto:true,
    cols:['Clave','Proyecto','Codigo','Descripcion','Ud','PrecioUnitario','CosteUnitario','Grupo'],
    num:['PrecioUnitario','CosteUnitario'] },
  M_MOExterna: { target:'moExterna', proyecto:true,
    cols:['Clave','Proyecto','CodigoPerfil','Descripcion','Empresa','Contrato','TarifaHora','Observaciones'],
    num:['TarifaHora'] },
  M_Alquileres: { target:'alquileres', proyecto:true,
    cols:['Clave','Proyecto','CodigoRef','Descripcion','Categoria','Empresa','Contrato','TarifaHora','TarifaDia'],
    num:['TarifaHora','TarifaDia'] },
  M_Indirectos: { target:'indirectos', proyecto:true,
    cols:['Clave','Proyecto','Codigo','Categoria','Concepto','ImporteEstimado','Observaciones'],
    num:['ImporteEstimado'] },
  Planificacion: { target:'planificacion', proyecto:true,
    cols:['Clave','Proyecto','Fecha','Semana','CodProduccion','Descripcion','Ud','CantidadPlanif','VentaPlanif','CostePlanif'],
    num:['Semana','CantidadPlanif','VentaPlanif','CostePlanif'] }
};

// Orden de lectura (corporativos primero)
var ORDEN = ['M_Proyectos','M_Personal','M_Maquinaria','M_Partidas','M_PartidasControl',
             'M_Subcontratas','M_Materiales','M_MOExterna','M_Alquileres','M_Indirectos','Planificacion'];

function internalName(i) { return i === 0 ? 'Title' : 'field_' + i; }

function num(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

// Mapea una fila cruda de SP -> objeto con nombres reales de columna.
// Usa fieldMap { TituloColumna: InternalNameReal } para leer POR NOMBRE.
// La columna 0 (Clave) SIEMPRE se lee de 'Title': es el campo de título nativo
// de SharePoint y donde reside el valor real (su columna "Clave"/LinkTitle es
// calculada y no devuelve valor en la lectura de items).
// Para el resto, si fieldMap no trae una columna, cae al nombre posicional
// field_N como respaldo (compatibilidad con listas antiguas incompletas).
function mapRow(raw, schema, fieldMap) {
  var o = {}, cols = schema.cols, numset = {};
  (schema.num || []).forEach(function (k) { numset[k] = 1; });
  for (var i = 0; i < cols.length; i++) {
    var name = cols[i];
    var internal;
    if (i === 0) {
      internal = 'Title';                                   // Clave: siempre Title
    } else {
      internal = (fieldMap && fieldMap[name]) ? fieldMap[name] : internalName(i);
    }
    var val = raw[internal];
    if (val === undefined) val = null;
    o[name] = numset[name] ? num(val) : val;   // resto se deja tal cual (texto/código/fecha)
  }
  return o;
}

// ¿Fila vacía? (sin Clave; descarta filas de prueba en blanco)
function filaVacia(o) {
  var clave = o.Clave;
  return clave === null || clave === undefined || String(clave).trim() === '';
}

/* ---------- LECTURA COMPLETA ---------- */
// Devuelve { proyectos:[...], personal:[...], ... } con nombres reales.
// onProgress(msg) opcional. Por defecto descarta filas sin Clave (pruebas).
async function leerTodo(onProgress, incluirVacias) {
  await ensureReady();
  var log = onProgress || function(){};
  var DATOS = {}, META = {};
  for (var i = 0; i < ORDEN.length; i++) {
    var listName = ORDEN[i];
    var sch = SP_SCHEMA[listName];
    log('Leyendo ' + listName + '…');
    var fieldMap = await spFieldMap(listName);   // nombre de columna -> InternalName real
    var rows = await spReadAll(listName);
    var mapped = [];
    var vacias = 0;
    rows.forEach(function (raw) {
      var o = mapRow(raw, sch, fieldMap);
      if (!incluirVacias && filaVacia(o)) { vacias++; return; }
      mapped.push(o);
    });
    DATOS[sch.target] = mapped;
    META[sch.target] = { lista:listName, total:rows.length, validas:mapped.length, vacias:vacias };
    log('  ' + listName + ': ' + mapped.length + '/' + rows.length + ' filas'
        + (vacias ? ' (' + vacias + ' vacías descartadas)' : ''));
  }
  DATOS._meta = META;
  return DATOS;
}

/* ---------- API pública ---------- */
global.SPTelice = {
  ensureReady: ensureReady,
  login: login,
  currentAccount: currentAccount,
  leerTodo: leerTodo,
  spGet: spGet,
  SP_SCHEMA: SP_SCHEMA
};

})(window);
