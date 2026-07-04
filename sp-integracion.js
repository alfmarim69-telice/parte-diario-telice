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
    num:['PrecioUnitario','CosteUnitario'],
    // Nombres internos reales (SharePoint codifica ó → _x00f3_)
    _fieldMap: {
      'Clave':          'Title',
      'Proyecto':       'Proyecto',
      'Codigo':         'C_x00f3_digo',
      'Descripcion':    'Descripci_x00f3_n',
      'Ud':             'Ud',
      'PrecioUnitario': 'field_5',
      'CosteUnitario':  'field_6',
      'Grupo':          'Grupo'
    }
  },
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
    cols:['Clave','Proyecto','Cuadrilla','Fecha','Semana','CodProduccion','Descripcion','Ud','CantidadPlanif','VentaPlanif','CostePlanif','PKs'],
    num:['Semana','CantidadPlanif','VentaPlanif','CostePlanif'],
    _fieldMap:{
      'Clave':          'Title',
      'Proyecto':       'field_1',
      'Fecha':          'Fecha',
      'Semana':         'field_3',
      'CodProduccion':  'field_4',
      'Descripcion':    'field_5',
      'Ud':             'field_6',
      'CantidadPlanif': 'field_7',
      'VentaPlanif':    'field_8',
      'CostePlanif':    'field_9',
      'Cuadrilla':      'Cuadrilla',
      'PKs':            'PKs'
    }
  },
  ParteMateriales: { target:'parteMateriales', proyecto:true,
    cols:['Clave','ParteId','Proyecto','CodMaterial','CodControl','Cantidad','TipoMovimiento','CosteUnitario'],
    num:['CodControl','Cantidad','CosteUnitario'],
    _fieldMap:{
      'Clave':'Title','ParteId':'field_1','Proyecto':'field_2',
      'CodMaterial':'field_3','CodControl':'field_4','Cantidad':'field_5',
      'TipoMovimiento':'TipoMovimiento','CosteUnitario':'CosteUnitario'
    }
  }
};

// Orden de lectura (corporativos primero)
var ORDEN = ['M_Proyectos','M_Personal','M_Maquinaria','M_Partidas','M_PartidasControl',
             'M_Subcontratas','M_Materiales','M_MOExterna','M_Alquileres','M_Indirectos',
             'Planificacion','ParteMateriales'];

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
    var fieldMap = sch._fieldMap || await spFieldMap(listName);   // usar mapa fijo si existe
    var rows = await spReadAll(listName);
    var mapped = [];
    var vacias = 0;
    rows.forEach(function (raw) {
      var o = mapRow(raw, sch, fieldMap);
      if (!incluirVacias && filaVacia(o)) { vacias++; return; }
      o._sp_id = raw.Id || raw.ID || null;   // ID nativo de SharePoint
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

/* ---------- ESCRITURA ---------- */
async function spPost(listTitle, body) {
  var t = await getToken();
  var r = await fetch(spUrl("/lists/GetByTitle('" + listTitle + "')/items"), {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + t,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata'
    },
    body: JSON.stringify(body)
  });
  if (!r.ok) {
    var msg = 'HTTP ' + r.status;
    try {
      var j = await r.json();
      var em = j.error && j.error.message;
      msg += ': ' + (typeof em === 'string' ? em : (em && em.value) || JSON.stringify(j));
    } catch(e) {}
    throw new Error(msg);
  }
  return r.json();
}

async function spPatch(listTitle, id, body) {
  var t = await getToken();
  var r = await fetch(spUrl("/lists/GetByTitle('" + listTitle + "')/items(" + id + ")"), {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + t,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    },
    body: JSON.stringify(body)
  });
  if (!r.ok && r.status !== 204) {
    var msg = 'HTTP ' + r.status;
    try {
      var j = await r.json();
      var em = j.error && j.error.message;
      msg += ': ' + (typeof em === 'string' ? em : (em && em.value) || JSON.stringify(j));
    } catch(e) {}
    throw new Error(msg);
  }
  return true;
}

async function spDelete(listTitle, id) {
  var t = await getToken();
  var r = await fetch(spUrl("/lists/GetByTitle('" + listTitle + "')/items(" + id + ")"), {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + t,
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    }
  });
  if (!r.ok && r.status !== 204) throw new Error('HTTP ' + r.status);
  return true;
}

// Mapa fijo de ParteMateriales
var PARTEMAT_FIELD_MAP = {
  'Clave':'Title','ParteId':'field_1','Proyecto':'field_2',
  'CodMaterial':'field_3','CodControl':'field_4','Cantidad':'field_5',
  'TipoMovimiento':'TipoMovimiento','CosteUnitario':'CosteUnitario'
};

// Registra un movimiento de material (ENTRADA o SALIDA) en SharePoint
async function registrarMovimientoMaterial(mov) {
  // mov: { parteId, proyecto, codMaterial, codControl, cantidad, tipo, costeUnitario }
  var clave = mov.proyecto + '|' + mov.parteId + '|' + mov.codMaterial + '|' + mov.tipo;
  var body = { Title: clave };
  var campos = {
    'ParteId': mov.parteId || '',
    'Proyecto': mov.proyecto || '',
    'CodMaterial': mov.codMaterial || '',
    'CodControl': mov.codControl || 0,
    'Cantidad': mov.cantidad || 0,
    'TipoMovimiento': mov.tipo || 'ENTRADA',
    'CosteUnitario': mov.costeUnitario || 0
  };
  Object.keys(campos).forEach(function(col) {
    body[PARTEMAT_FIELD_MAP[col]] = campos[col];
  });
  return spPost('ParteMateriales', body);
}

// Obtiene el stock disponible de un material en un proyecto
// Devuelve { entradas, salidas, disponible }
async function getStockMaterial(proyecto, codMaterial) {
  var r = await spGet(
    "/lists/GetByTitle('ParteMateriales')/items" +
    "?$filter=field_2 eq '" + proyecto + "' and field_3 eq '" + codMaterial + "'" +
    "&$select=field_5,TipoMovimiento&$top=5000"
  );
  var entradas = 0, salidas = 0;
  (r.value || []).forEach(function(item) {
    var cant = item.field_5 || 0;
    if (item.TipoMovimiento === 'SALIDA') salidas += cant;
    else entradas += cant;
  });
  return { entradas: entradas, salidas: salidas, disponible: entradas - salidas };
}

// Obtiene stock de todos los materiales de un proyecto
async function getStockProyecto(proyecto) {
  var r = await spGet(
    "/lists/GetByTitle('ParteMateriales')/items" +
    "?$filter=field_2 eq '" + proyecto + "'" +
    "&$select=field_3,field_5,TipoMovimiento&$top=5000"
  );
  var stock = {};
  (r.value || []).forEach(function(item) {
    var cod = item.field_3 || '';
    if (!stock[cod]) stock[cod] = { entradas: 0, salidas: 0, disponible: 0 };
    var cant = item.field_5 || 0;
    if (item.TipoMovimiento === 'SALIDA') stock[cod].salidas += cant;
    else stock[cod].entradas += cant;
    stock[cod].disponible = stock[cod].entradas - stock[cod].salidas;
  });
  return stock;
}
var MAT_FIELD_MAP = {
  'Proyecto':       'Proyecto',
  'Codigo':         'C_x00f3_digo',
  'Descripcion':    'Descripci_x00f3_n',
  'Ud':             'Ud',
  'PrecioUnitario': 'field_5',
  'CosteUnitario':  'field_6',
  'Grupo':          'Grupo'
};

var MAT_CAMPOS_DATOS = ['Proyecto','Codigo','Descripcion','Ud','PrecioUnitario','CosteUnitario','Grupo'];

// Construye body para spPost usando mapa fijo (sin depender de spFieldMap)
function buildMatBody(clave, campos) {
  var body = { Title: clave };
  MAT_CAMPOS_DATOS.forEach(function(col) {
    if (campos[col] === undefined) return;
    body[MAT_FIELD_MAP[col]] = campos[col];
  });
  return body;
}

// Puebla M_Materiales desde M_Partidas si está vacía para el proyecto dado.
async function inicializarMateriales(proyecto, partidas) {
  var todos = await spGet(
    "/lists/GetByTitle('M_Materiales')/items?$top=5000&$select=Id,Proyecto"
  );
  var existentes = (todos.value || []).filter(function(x) {
    return (x.Proyecto || '') === proyecto;
  });
  if (existentes.length > 0) {
    return { creados: 0, yaExistian: true };
  }

  var matPartidas = (partidas || []).filter(function (p) {
    return (p.Proyecto || '').trim() === proyecto &&
           String(p.CtrlSuministro || '').trim().indexOf('2.') === 0;
  });

  var creados = 0;
  for (var i = 0; i < matPartidas.length; i++) {
    var p = matPartidas[i];
    var clave = proyecto + '|' + (p.Codigo || '');
    var body = buildMatBody(clave, {
      'Proyecto':       proyecto,
      'Codigo':         p.Codigo || '',
      'Descripcion':    p.Descripcion || '',
      'Ud':             p.Ud || '',
      'PrecioUnitario': p.PuVentaSum || 0,
      'CosteUnitario':  p.CosteUdSum || 0,
      'Grupo':          String(p.CtrlSuministro || '').trim()
    });
    await spPost('M_Materiales', body);
    creados++;
  }
  return { creados: creados, yaExistian: false };
}

// Añade un material nuevo a M_Materiales
async function addMaterial(proyecto, mat) {
  var clave = proyecto + '|' + mat.cod;
  var body = buildMatBody(clave, {
    'Proyecto':       proyecto,
    'Codigo':         mat.cod,
    'Descripcion':    mat.desc,
    'Ud':             mat.ud,
    'PrecioUnitario': mat.pu || 0,
    'CosteUnitario':  mat.coste_ud || 0,
    'Grupo':          mat.grupo
  });
  var item = await spPost('M_Materiales', body);
  return item.Id || item.id || null;
}

// Elimina un material de M_Materiales por su ID de SharePoint
async function deleteMaterial(spId) {
  return spDelete('M_Materiales', spId);
}


// ═══════════════════════════════════════════════════════
// TAREAS POR PARTIDA (despiece de la partida en tareas con peso %)
// Lista M_TareasPartida: Proyecto | Codigo | Orden | Tarea | Peso
// Title = Proyecto|Codigo|Orden. Los pesos de una partida suman 100.
// ═══════════════════════════════════════════════════════
var _tareasFieldMapCache = null;
async function getTareasFieldMap() {
  if (!_tareasFieldMapCache) _tareasFieldMapCache = await spFieldMap('M_TareasPartida');
  return _tareasFieldMapCache;
}

// Devuelve todas las tareas definidas de una obra:
// [{Codigo, Orden, Tarea, Peso}] ordenadas por Codigo y Orden
async function leerTareasPartida(proyecto) {
  await ensureReady();
  try {
    var fm = await getTareasFieldMap();
    var cP = fm['Proyecto'] || 'field_1', cC = fm['Codigo'] || 'field_2';
    var cO = fm['Orden'] || 'field_3', cT = fm['Tarea'] || 'field_4', cPe = fm['Peso'] || 'field_5';
    var data = await spGet("/lists/GetByTitle('M_TareasPartida')/items?$top=5000&$select=" +
      [cP, cC, cO, cT, cPe].join(','));
    var out = [];
    (data.value || []).forEach(function (it) {
      if (String(it[cP] || '').trim() !== proyecto) return;
      out.push({ Codigo: it[cC] || '', Orden: Number(it[cO]) || 0, Tarea: it[cT] || '', Peso: Number(it[cPe]) || 0 });
    });
    out.sort(function (a, b) { return a.Codigo === b.Codigo ? a.Orden - b.Orden : a.Codigo.localeCompare(b.Codigo); });
    return out;
  } catch (e) {
    console.warn('[SP] M_TareasPartida no disponible:', e && e.message);
    return [];
  }
}

// Sincroniza las tareas de UNA partida: crea/actualiza las pasadas y borra
// de SharePoint las que ya no estén. tareas: [{Orden, Tarea, Peso}]
async function guardarTareasPartida(proyecto, codigo, tareas) {
  await ensureReady();
  var fm = await getTareasFieldMap();
  var cP = fm['Proyecto'] || 'field_1';
  var cC = fm['Codigo'] || 'field_2';

  var existentes = await spGet("/lists/GetByTitle('M_TareasPartida')/items?$select=Id,Title," + cP + "," + cC + "&$top=5000");
  var porClave = {};
  (existentes.value || []).forEach(function (it) {
    if (String(it[cP] || '').trim() !== proyecto) return;
    if (String(it[cC] || '').trim() !== codigo) return;
    porClave[it.Title] = it.Id;
  });

  var errores = [], creadas = 0, actualizadas = 0, borradas = 0;
  var clavesNuevas = {};
  for (var i = 0; i < tareas.length; i++) {
    var t = tareas[i];
    var clave = proyecto + '|' + codigo + '|' + t.Orden;
    clavesNuevas[clave] = 1;
    var body = {};
    body[fm['Proyecto'] || 'field_1'] = proyecto;
    body[fm['Codigo'] || 'field_2'] = codigo;
    body[fm['Orden'] || 'field_3'] = Number(t.Orden) || 0;
    body[fm['Tarea'] || 'field_4'] = t.Tarea || '';
    body[fm['Peso'] || 'field_5'] = Number(t.Peso) || 0;
    try {
      if (porClave[clave]) { await spPatch('M_TareasPartida', porClave[clave], body); actualizadas++; }
      else { body.Title = clave; await spPost('M_TareasPartida', body); creadas++; }
    } catch (e) { errores.push(clave + ': ' + e.message); }
  }
  var aBorrar = Object.keys(porClave).filter(function (cl) { return !clavesNuevas[cl]; });
  for (var j = 0; j < aBorrar.length; j++) {
    try { await spDelete('M_TareasPartida', porClave[aBorrar[j]]); borradas++; }
    catch (e) { errores.push('borrar ' + aBorrar[j] + ': ' + e.message); }
  }
  return { creadas: creadas, actualizadas: actualizadas, borradas: borradas, errores: errores, ok: errores.length === 0 };
}

// Acumulado ejecutado por tarea en una obra (de ParteProduccionTareas):
// devuelve { cod: { tarea: cantidadAcumulada } }. Para pintar en el parte
// el ejecutado/pendiente de cada tarea, al estilo ficha de elemento.
async function leerProduccionTareas(proyecto) {
  await ensureReady();
  try {
    var data = await spGet("/lists/GetByTitle('ParteProduccionTareas')/items?$top=5000&$select=field_2,field_3,field_4,field_5");
    var out = {};
    (data.value || []).forEach(function (it) {
      if (String(it.field_2 || '').trim() !== proyecto) return;
      var cod = it.field_3 || '', tarea = it.field_4 || '';
      if (!cod || !tarea) return;
      if (!out[cod]) out[cod] = {};
      out[cod][tarea] = (out[cod][tarea] || 0) + (Number(it.field_5) || 0);
    });
    return out;
  } catch (e) {
    console.warn('[SP] ParteProduccionTareas no disponible:', e && e.message);
    return {};
  }
}

// Carga masiva de despiece de tareas para MUCHAS partidas a la vez (pensado
// para volcar de golpe un Excel con cientos de filas, como el propuesto por
// IA y revisado por el usuario). A diferencia de guardarTareasPartida (que
// lee la lista completa cada vez), aquí se lee M_TareasPartida UNA sola vez
// y se reparte en memoria — imprescindible para no hacer 295 lecturas.
// El borrado de tareas obsoletas solo afecta a las partidas presentes en
// `filas`; una partida que no aparezca en esta carga no se toca.
// filas: [{Proyecto, Codigo, Orden, Tarea, Peso}]
async function guardarDespieceMasivo(filas, onProgreso) {
  await ensureReady();
  var fm = await getTareasFieldMap();
  var cP = fm['Proyecto'] || 'field_1', cC = fm['Codigo'] || 'field_2';

  var existentes = await spGet("/lists/GetByTitle('M_TareasPartida')/items?$select=Id,Title," + cP + "," + cC + "&$top=5000");
  var porClave = {};
  (existentes.value || []).forEach(function (it) { porClave[it.Title] = it.Id; });

  // Agrupar filas de entrada por partida
  var grupos = {};
  filas.forEach(function (f) {
    var key = f.Proyecto + '|' + f.Codigo;
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push({ Orden: Number(f.Orden) || 0, Tarea: f.Tarea || '', Peso: Number(f.Peso) || 0 });
  });

  var clavesNuevasGlobal = {};
  Object.keys(grupos).forEach(function (key) {
    var partes = key.split('|'), proyecto = partes[0], codigo = partes[1];
    grupos[key].forEach(function (t) { clavesNuevasGlobal[proyecto + '|' + codigo + '|' + t.Orden] = 1; });
  });

  // Tareas a borrar: existían, pertenecen a una partida incluida en esta
  // carga, pero ya no están en la nueva versión de esa partida
  var aBorrarIds = [];
  (existentes.value || []).forEach(function (it) {
    var key = String(it[cP] || '').trim() + '|' + String(it[cC] || '').trim();
    if (!grupos[key]) return;
    if (!clavesNuevasGlobal[it.Title]) aBorrarIds.push(it.Id);
  });

  var errores = [], creadas = 0, actualizadas = 0, borradas = 0;
  var claves = Object.keys(grupos), total = claves.length, i = 0;
  for (var k = 0; k < claves.length; k++) {
    var key = claves[k], partes2 = key.split('|'), proyecto2 = partes2[0], codigo2 = partes2[1];
    var tareas = grupos[key].slice().sort(function (a, b) { return a.Orden - b.Orden; });
    i++;
    if (onProgreso) onProgreso(i, total, proyecto2, codigo2);
    for (var j = 0; j < tareas.length; j++) {
      var t = tareas[j];
      var clave = proyecto2 + '|' + codigo2 + '|' + t.Orden;
      var body = {};
      body[fm['Proyecto'] || 'field_1'] = proyecto2;
      body[fm['Codigo'] || 'field_2'] = codigo2;
      body[fm['Orden'] || 'field_3'] = t.Orden;
      body[fm['Tarea'] || 'field_4'] = t.Tarea;
      body[fm['Peso'] || 'field_5'] = t.Peso;
      try {
        if (porClave[clave]) { await spPatch('M_TareasPartida', porClave[clave], body); actualizadas++; }
        else { body.Title = clave; await spPost('M_TareasPartida', body); creadas++; }
      } catch (e) { errores.push(clave + ': ' + e.message); }
    }
  }
  for (var b = 0; b < aBorrarIds.length; b++) {
    try { await spDelete('M_TareasPartida', aBorrarIds[b]); borradas++; }
    catch (e) { errores.push('borrar id ' + aBorrarIds[b] + ': ' + e.message); }
  }

  return { partidas: total, creadas: creadas, actualizadas: actualizadas, borradas: borradas, errores: errores, ok: errores.length === 0 };
}

// ═══════════════════════════════════════════════════════
// GUARDADO DE PLANIFICACIÓN (comparación planificado vs ejecutado)
// ═══════════════════════════════════════════════════════

// El mapa de nombres internos se resuelve una sola vez por nombre real de
// columna (igual que en la lectura), así no dependemos de que la lista se
// haya creado con field_1, field_2... — funciona la haya creado quien la haya
// creado (UI de SharePoint o API).
var _planifFieldMapCache = null;  // se resetea aquí: null fuerza recarga en cada nueva sesión
async function getPlanifFieldMap() {
  if (!_planifFieldMapCache) {
    _planifFieldMapCache = await spFieldMap('Planificacion');
  }
  return _planifFieldMapCache;
}

// Guarda el plan completo de una obra+cuadrilla: actualiza lo que ya existía,
// crea lo nuevo, y BORRA de SharePoint las líneas de esa obra+cuadrilla que ya
// no están en `filas`. `filas` debe traer el estado COMPLETO deseado para esa
// obra+cuadrilla, no solo lo cambiado. El resto de cuadrillas no se toca.
// filas: [{ Clave, Proyecto, Cuadrilla, Fecha, Semana, CodProduccion, Descripcion,
//           Ud, CantidadPlanif, VentaPlanif, CostePlanif, PKs }]
async function guardarPlanificacion(proyecto, cuadrilla, filas) {
  await ensureReady();
  // Usar el mapa fijo del esquema (más fiable que spFieldMap dinámico)
  var fm = (SP_SCHEMA.Planificacion && SP_SCHEMA.Planificacion._fieldMap) || await getPlanifFieldMap();
  var colProyecto = fm['Proyecto'] || 'field_1';
  var colCuadrilla = fm['Cuadrilla'] || 'Cuadrilla';

  // Lectura completa + filtrado en cliente ($filter sobre columnas no
  // indexadas devuelve HTTP 500 en estas listas)
  var existentes = await spGet(
    "/lists/GetByTitle('Planificacion')/items?$select=Id,Title," + colProyecto + "," + colCuadrilla + "&$top=5000"
  );
  var porClave = {};
  (existentes.value || []).forEach(function (it) {
    if (String(it[colProyecto] || '').trim() !== proyecto) return;
    if (String(it[colCuadrilla] || '').trim() !== cuadrilla) return;
    porClave[it.Title] = it.Id;
  });

  var errores = [], creadas = 0, actualizadas = 0, borradas = 0;

  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    var campos = {
      Proyecto: f.Proyecto, Cuadrilla: f.Cuadrilla || cuadrilla,
      Fecha: f.Fecha, Semana: parseInt(f.Semana) || 0,
      CodProduccion: f.CodProduccion, Descripcion: f.Descripcion, Ud: f.Ud,
      CantidadPlanif: f.CantidadPlanif || 0,
      VentaPlanif: f.VentaPlanif || 0,
      CostePlanif: f.CostePlanif || 0
    };
    if (f.PKs) campos.PKs = f.PKs;
    var body = {};
    Object.keys(campos).forEach(function (col) { body[fm[col] || col] = campos[col]; });
    try {
      if (porClave[f.Clave]) {
        await spPatch('Planificacion', porClave[f.Clave], body);
        actualizadas++;
      } else {
        body.Title = f.Clave;
        await spPost('Planificacion', body);
        creadas++;
      }
    } catch (e) { errores.push(f.Clave + ': ' + e.message); }
  }

  var clavesNuevas = {};
  filas.forEach(function (f) { clavesNuevas[f.Clave] = 1; });
  var aBorrar = Object.keys(porClave).filter(function (cl) { return !clavesNuevas[cl]; });
  for (var j = 0; j < aBorrar.length; j++) {
    try { await spDelete('Planificacion', porClave[aBorrar[j]]); borradas++; }
    catch (e) { errores.push('borrar ' + aBorrar[j] + ': ' + e.message); }
  }

  return { creadas: creadas, actualizadas: actualizadas, borradas: borradas, errores: errores, ok: errores.length === 0 };
}

// Lee la producción REAL de una obra desde ParteProduccion, cruzándola con
// Partes para obtener fecha y cuadrilla (ParteProduccion no las tiene: viven
// en la cabecera del parte y se unen por ParteId).
// Filtrado por proyecto en cliente ($filter sobre columnas no indexadas da HTTP 500).
// Devuelve [{ ParteId, Fecha, Cuadrilla, CodProduccion, Cantidad, Importe, Ud, PKInicio, PKFinal }]
async function leerProduccionReal(proyecto) {
  await ensureReady();

  // 1) Cabeceras: field_1=ParteId, field_2=Proyecto, field_3=Fecha, field_6=Cuadrilla
  var partes = await spGet("/lists/GetByTitle('Partes')/items?$top=5000&$select=field_1,field_2,field_3,field_6");
  var cabPorParte = {};
  (partes.value || []).forEach(function (it) {
    if (String(it.field_2 || '').trim() !== proyecto) return;
    cabPorParte[String(it.field_1)] = { fecha: it.field_3 || '', cuadrilla: (it.field_6 || '').trim() };
  });

  // 2) Líneas de producción (field_1=ParteId, field_2=Proyecto, field_3=Cod, field_4=Cantidad)
  var prod = await spGet("/lists/GetByTitle('ParteProduccion')/items?$top=5000&$select=field_1,field_2,field_3,field_4,Importe,Ud,PKInicio,PKFinal");
  var out = [];
  (prod.value || []).forEach(function (it) {
    if (String(it.field_2 || '').trim() !== proyecto) return;
    var cab = cabPorParte[String(it.field_1)] || {};
    out.push({
      ParteId: it.field_1,
      Fecha: cab.fecha || '',
      Cuadrilla: cab.cuadrilla || '',
      CodProduccion: it.field_3 || '',
      Cantidad: it.field_4 || 0,
      Importe: it.Importe || 0,
      Ud: it.Ud || '',
      PKInicio: it.PKInicio || '',
      PKFinal: it.PKFinal || ''
    });
  });
  return out;
}

// Cuadrillas distintas ya usadas en los partes de una obra (para el selector
// de la planificación: se planifica POR CUADRILLA, no por obra entera).
async function leerCuadrillas(proyecto) {
  await ensureReady();
  var partes = await spGet("/lists/GetByTitle('Partes')/items?$top=5000&$select=field_2,field_6");
  var set = {};
  (partes.value || []).forEach(function (it) {
    if (String(it.field_2 || '').trim() !== proyecto) return;
    var c = (it.field_6 || '').trim();
    if (c) set[c] = 1;
  });
  return Object.keys(set).sort();
}

// Medición por kilómetro desde la lista M_MedicionPK
// (Proyecto | Codigo | Km | Medicion; Title = Proyecto|Codigo|Km).
// Filtrado en cliente. Si la lista no existe aún, devuelve [] sin romper.
var _medPKFieldMapCache = null;
async function leerMedicionPK(proyecto) {
  await ensureReady();
  try {
    if (!_medPKFieldMapCache) _medPKFieldMapCache = await spFieldMap('M_MedicionPK');
    var fm = _medPKFieldMapCache;
    var cP = fm['Proyecto'] || 'field_1', cC = fm['Codigo'] || 'field_2';
    var cK = fm['Km'] || 'field_3', cM = fm['Medicion'] || 'field_4';
    var data = await spGet("/lists/GetByTitle('M_MedicionPK')/items?$top=5000&$select=" +
      [cP, cC, cK, cM].join(','));
    var out = [];
    (data.value || []).forEach(function (it) {
      if (String(it[cP] || '').trim() !== proyecto) return;
      out.push({ Codigo: it[cC] || '', Km: Number(it[cK]) || 0, Medicion: Number(it[cM]) || 0 });
    });
    return out;
  } catch (e) {
    console.warn('[SP] M_MedicionPK no disponible:', e && e.message);
    return [];
  }
}

// Carga masiva de mediciones por PK (para la página carga-mediciones.html).
// filas: [{Proyecto, Codigo, Km, Medicion}]. Upsert por Title=Proyecto|Codigo|Km.
async function guardarMedicionPK(filas) {
  await ensureReady();
  if (!_medPKFieldMapCache) _medPKFieldMapCache = await spFieldMap('M_MedicionPK');
  var fm = _medPKFieldMapCache;
  var existentes = await spGet("/lists/GetByTitle('M_MedicionPK')/items?$select=Id,Title&$top=5000");
  var porClave = {};
  (existentes.value || []).forEach(function (it) { porClave[it.Title] = it.Id; });

  var errores = [], creadas = 0, actualizadas = 0;
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    var clave = f.Proyecto + '|' + f.Codigo + '|' + f.Km;
    var body = {};
    body[fm['Proyecto'] || 'field_1'] = f.Proyecto;
    body[fm['Codigo'] || 'field_2'] = f.Codigo;
    body[fm['Km'] || 'field_3'] = Number(f.Km) || 0;
    body[fm['Medicion'] || 'field_4'] = Number(f.Medicion) || 0;
    try {
      if (porClave[clave]) { await spPatch('M_MedicionPK', porClave[clave], body); actualizadas++; }
      else { body.Title = clave; await spPost('M_MedicionPK', body); creadas++; }
    } catch (e) { errores.push(clave + ': ' + e.message); }
  }
  return { creadas: creadas, actualizadas: actualizadas, errores: errores, ok: errores.length === 0 };
}

// ═══════════════════════════════════════════════════════
// GUARDADO COMPLETO DEL PARTE EN SHAREPOINT
// ═══════════════════════════════════════════════════════

// Mapas de nombres internos para cada lista de partes
var PARTE_MAPS = {
  Partes: {
    ParteId:'field_1', Proyecto:'field_2', Fecha:'field_3', Semana:'field_4',
    Responsable:'field_5', Cuadrilla:'field_6', UnidadProductiva:'field_7',
    Estado:'field_8', Meteorologia:'field_9', Lugar:'field_10', Observaciones:'field_11',
    HoraEntrada:'HoraEntrada', HoraParada:'HoraParada', HoraReinicio:'HoraReinicio',
    HoraSalida:'HoraSalida', InicioFranja:'InicioFranja', FinFranja:'FinFranja',
    NumAutorizacion:'NumAutorizacion', Provincia:'Provincia', PKLugar:'PKLugar'
  },
  PartePersonal: {
    ParteId:'field_1', Proyecto:'field_2', NumPersonal:'field_3', CodControl:'field_4',
    Horas:'field_5', Dieta:'field_6', Km:'field_7',
    CosteJornada:'CosteJornada', Tipo:'Tipo', Nombre:'Nombre'
  },
  ParteMaquinaria: {
    ParteId:'field_1', Proyecto:'field_2', CodMaquina:'field_3', CodControl:'field_4',
    Horas:'field_5', CosteHora:'CosteHora', CosteTotal:'CosteTotal', Proveedor:'Proveedor'
  },
  ParteMOExterna: {
    ParteId:'field_1', Proyecto:'field_2', CodPerfil:'field_3', CodControl:'field_4',
    Horas:'field_5', Importe:'field_6', Empresa:'Empresa', TarifaHora:'TarifaHora'
  },
  ParteOtros: {
    ParteId:'field_1', Proyecto:'field_2', CodControl:'field_3', Concepto:'field_4',
    Importe:'field_5', Observaciones:'field_6'
  },
  ParteSubcontratas: {
    ParteId:'field_1', Proyecto:'field_2', CodControl:'field_3', Empresa:'field_4',
    Cantidad:'field_5', Importe:'field_6',
    CodPartida:'CodPartida', Ud:'Ud', Tarifa:'Tarifa',
    PKInicio:'PKInicio', PKFinal:'PKFinal', Contrato:'Contrato'
  },
  ParteProduccion: {
    ParteId:'field_1', Proyecto:'field_2', CodProduccion:'field_3', Cantidad:'field_4',
    Observaciones:'field_5', Importe:'Importe', Ud:'Ud',
    PKInicio:'PKInicio', PKFinal:'PKFinal', Lugar:'Lugar', TaskBC:'TaskBC'
  },
  ParteProduccionTareas: {
    ParteId:'field_1', Proyecto:'field_2', CodProduccion:'field_3', Tarea:'field_4',
    Cantidad:'field_5', PKInicio:'field_6', PKFinal:'field_7'
  },
  ParteAlquileres: {
    ParteId:'field_1', Proyecto:'field_2', CodRef:'field_3', CodControl:'field_4',
    Horas:'field_5', Importe:'field_6', Descripcion:'Descripcion', CosteHora:'CosteHora'
  }
};

function buildParteBody(listName, clave, campos) {
  var fm = PARTE_MAPS[listName] || {};
  var body = { Title: clave };
  Object.keys(campos).forEach(function(col) {
    var internal = fm[col] || col;
    if (campos[col] !== undefined && campos[col] !== null) {
      body[internal] = campos[col];
    }
  });
  return body;
}

async function guardarParteEnSP(parte) {
  var parteId = parte._id ? String(parte._id) : String(Date.now());
  var proy    = parte.cabecera.proyecto || '';
  var clave   = proy + '|' + parteId;
  var errores = [];

  // 1) Cabecera
  try {
    var cab = parte.cabecera;
    var body = buildParteBody('Partes', clave, {
      ParteId: parteId, Proyecto: proy,
      Fecha: cab.fecha || '', Semana: cab.semana || 0,
      Responsable: cab.responsable || '', Cuadrilla: cab.cuadrilla || '',
      UnidadProductiva: cab.unidad_prod || '', Estado: cab.estado || 'Enviado',
      Meteorologia: cab.meteo || '', Lugar: cab.lugar || '',
      Observaciones: cab.obs || '',
      HoraEntrada: cab.hora_entrada || '', HoraParada: cab.hora_parada || '',
      HoraReinicio: cab.hora_reinicio || '', HoraSalida: cab.hora_salida || '',
      InicioFranja: cab.franja_ini || '', FinFranja: cab.franja_fin || '',
      NumAutorizacion: cab.num_autorizacion || '',
      Provincia: cab.provincia || '', PKLugar: cab.pk_lugar || ''
    });
    await spPost('Partes', body);
  } catch(e) { errores.push('Partes: ' + e.message); }

  // 2) Personal
  for (var i = 0; i < (parte.personal || []).length; i++) {
    var p = parte.personal[i];
    try {
      await spPost('PartePersonal', buildParteBody('PartePersonal',
        clave + '|P' + i, {
          ParteId: parteId, Proyecto: proy,
          NumPersonal: p.num || p.id || 0,
          CodControl: p.task_cod || 0,
          Horas: p.horas || 0, Dieta: p.dieta || 'sin_dieta',
          Km: p.km || 0, CosteJornada: p.coste_jornada || 0,
          Tipo: p.tipo || '', Nombre: p.nombre || ''
        }));
    } catch(e) { errores.push('Personal[' + i + ']: ' + e.message); }
  }

  // 3) Maquinaria
  for (var i = 0; i < (parte.maquinaria || []).length; i++) {
    var m = parte.maquinaria[i];
    try {
      await spPost('ParteMaquinaria', buildParteBody('ParteMaquinaria',
        clave + '|M' + i, {
          ParteId: parteId, Proyecto: proy,
          CodMaquina: m.cod || '', CodControl: m.task_cod || 0,
          Horas: m.horas || 0, CosteHora: m.ch || 0,
          CosteTotal: m.coste_total || 0, Proveedor: m.prov || ''
        }));
    } catch(e) { errores.push('Maquinaria[' + i + ']: ' + e.message); }
  }

  // 4) MO Externa
  for (var i = 0; i < (parte.mo_externa || []).length; i++) {
    var m = parte.mo_externa[i];
    try {
      await spPost('ParteMOExterna', buildParteBody('ParteMOExterna',
        clave + '|MOE' + i, {
          ParteId: parteId, Proyecto: proy,
          CodPerfil: m.cod || '', CodControl: m.task_cod || 0,
          Horas: m.horas || 0, Importe: m.importe || 0,
          Empresa: m.empresa || '', TarifaHora: m.tarifa || 0
        }));
    } catch(e) { errores.push('MOExterna[' + i + ']: ' + e.message); }
  }

  // 5) Subcontratas
  for (var i = 0; i < (parte.subcontratas || []).length; i++) {
    var s = parte.subcontratas[i];
    try {
      await spPost('ParteSubcontratas', buildParteBody('ParteSubcontratas',
        clave + '|S' + i, {
          ParteId: parteId, Proyecto: proy,
          CodControl: s.task || 0, Empresa: s.empresa || '',
          Cantidad: s.cant || 0, Importe: s.coste_total || 0,
          CodPartida: s.cod || '', Ud: s.ud || '',
          Tarifa: s.tarifa || 0, PKInicio: s.pk_ini || '',
          PKFinal: s.pk_fin || '', Contrato: s.contrato || ''
        }));
    } catch(e) { errores.push('Subcontratas[' + i + ']: ' + e.message); }
  }

  // 6) Producción
  for (var i = 0; i < (parte.produccion || []).length; i++) {
    var pr = parte.produccion[i];
    try {
      await spPost('ParteProduccion', buildParteBody('ParteProduccion',
        clave + '|PR' + i, {
          ParteId: parteId, Proyecto: proy,
          CodProduccion: pr.partida_cod || '',
          Cantidad: pr.cant || 0, Importe: pr.importe || 0,
          Ud: pr.ud || '', PKInicio: pr.pk_ini || '',
          PKFinal: pr.pk_fin || '', Lugar: pr.lugar || '',
          TaskBC: pr.task_grupo || '',
          Observaciones: pr.tareas && pr.tareas.length ? 'Equivalente por tareas' : ''
        }));
      // 6b) Detalle por tareas (si la partida tiene despiece):
      // la Cantidad de ParteProduccion es el EQUIVALENTE (Σ cant × peso);
      // aquí se guarda el desglose real que reportó el jefe de equipo.
      if (pr.tareas && pr.tareas.length) {
        for (var j = 0; j < pr.tareas.length; j++) {
          var ta = pr.tareas[j];
          if (!ta.cant || ta.cant <= 0) continue;
          await spPost('ParteProduccionTareas', buildParteBody('ParteProduccionTareas',
            clave + '|PR' + i + '|T' + j, {
              ParteId: parteId, Proyecto: proy,
              CodProduccion: pr.partida_cod || '',
              Tarea: ta.tarea || '', Cantidad: ta.cant || 0,
              PKInicio: pr.pk_ini || '', PKFinal: pr.pk_fin || ''
            }));
        }
      }
    } catch(e) { errores.push('Produccion[' + i + ']: ' + e.message); }
  }

  // 7) Alquileres
  for (var i = 0; i < (parte.alquileres || []).length; i++) {
    var a = parte.alquileres[i];
    try {
      await spPost('ParteAlquileres', buildParteBody('ParteAlquileres',
        clave + '|A' + i, {
          ParteId: parteId, Proyecto: proy,
          CodRef: a.cod || '', CodControl: a.task_cod || 0,
          Horas: a.horas || 0, Importe: a.coste_total || 0,
          Descripcion: a.desc || '', CosteHora: a.ch || 0
        }));
    } catch(e) { errores.push('Alquileres[' + i + ']: ' + e.message); }
  }

  // 8) Otros conceptos
  for (var i = 0; i < (parte.otros || []).length; i++) {
    var o = parte.otros[i];
    try {
      await spPost('ParteOtros', buildParteBody('ParteOtros',
        clave + '|O' + i, {
          ParteId: parteId, Proyecto: proy,
          CodControl: o.task_cod || 0,
          Concepto: o.concepto || o.obs || '',
          Importe: o.importe || o.coste_otros || 0,
          Observaciones: o.obs || ''
        }));
    } catch(e) { errores.push('Otros[' + i + ']: ' + e.message); }
  }

  return { parteId, errores, ok: errores.length === 0 };
}

/* ---------- API pública ---------- */
global.SPTelice = {
  ensureReady: ensureReady,
  login: login,
  currentAccount: currentAccount,
  leerTodo: leerTodo,
  spGet: spGet,
  spPost: spPost,
  spDelete: spDelete,
  SP_SCHEMA: SP_SCHEMA,
  inicializarMateriales: inicializarMateriales,
  addMaterial: addMaterial,
  deleteMaterial: deleteMaterial,
  registrarMovimientoMaterial: registrarMovimientoMaterial,
  getStockMaterial: getStockMaterial,
  getStockProyecto: getStockProyecto,
  guardarParteEnSP: guardarParteEnSP,
  guardarPlanificacion: guardarPlanificacion,
  leerProduccionReal: leerProduccionReal,
  leerCuadrillas: leerCuadrillas,
  leerMedicionPK: leerMedicionPK,
  guardarMedicionPK: guardarMedicionPK,
  leerTareasPartida: leerTareasPartida,
  guardarTareasPartida: guardarTareasPartida,
  leerProduccionTareas: leerProduccionTareas,
  guardarDespieceMasivo: guardarDespieceMasivo
};

})(window);

