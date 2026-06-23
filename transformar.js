/* ============================================================
   transformar.js · Parte Diario Telice
   DATOS (11 maestros SharePoint) -> OBRAS (estructura de la app)
   ------------------------------------------------------------
   OPCIÓN A: SharePoint guarda los maestros crudos; la app deriva
   certificación/presupuesto y agrupaciones.
   Requiere SPTelice.leerTodo() (sp-integracion.js).
   ============================================================ */
(function (global) {
'use strict';

// Índice (proyecto|codigo) -> CosteEstudio, desde subcontratas.
function indiceCosteSub(subcontratas) {
  var idx = {};
  subcontratas.forEach(function (s) {
    var k = (s.Proyecto || '') + '|' + (s.CodigoPartida || '');
    if (idx[k] === undefined) idx[k] = s.CosteEstudio;
  });
  return idx;
}

// Deriva _grupo/_tipo_grupo y te/ts/ti/to a partir de los Ctrl*.
function clasificaPartida(p) {
  var ti = p.CtrlIngenieria || '', ts = p.CtrlSuministro || '',
      te = p.CtrlEjecucion  || '', to = p.CtrlOtros || '';
  var grupo, tipo;
  if (ti)      { grupo = ti; tipo = 'ING'; }
  else if (to) { grupo = to; tipo = 'OTROS'; }
  else if (ts) { grupo = ts; tipo = 'SUM'; }
  else         { grupo = te; tipo = 'EJEC'; }
  return { te: te, ts: ts, ti: ti, to: to, _grupo: String(grupo || ''), _tipo_grupo: tipo };
}

function nuevaObra(cab) {
  return {
    codigo: cab.codigo, nombre: cab.nombre, tipo: cab.tipo, incentivos: cab.incentivos,
    partidas: [], partidas_cert: [], partidas_ppto: [],
    personal: [], mat: [], sub: [], maq_obra: [],
    mo_externa: [], alquileres: [], indirectos: [], planificacion: []
  };
}

function tipoApp(t) {
  t = String(t || '').toLowerCase();
  if (t.indexOf('catenaria') >= 0) return 'catenaria';
  if (t.indexOf('instal') >= 0)    return 'instalaciones';
  if (t.indexOf('subest') >= 0)    return 'subestaciones';
  return t;
}

function transformar(DATOS, opts) {
  opts = opts || {};
  var OBRAS = {};

  // 1) Cabeceras
  (DATOS.proyectos || []).forEach(function (p) {
    var cod = String(p.Codigo || p.Clave || '').trim();
    if (!cod) return;
    OBRAS[cod] = nuevaObra({
      codigo: cod,
      nombre: cod + ' · ' + (p.Nombre || ''),
      tipo: tipoApp(p.Tipo),
      incentivos: p.Incentivos || 'ninguno'
    });
  });

  var costeSub = indiceCosteSub(DATOS.subcontratas || []);

  // 2) Partidas (base operativa)
  (DATOS.partidas || []).forEach(function (p) {
    var obra = OBRAS[(p.Proyecto || '').trim()];
    if (!obra) return;
    var c = clasificaPartida(p);
    var k = (p.Proyecto || '') + '|' + (p.Codigo || '');
    obra.partidas.push({
      cod: p.Codigo || '', desc: p.Descripcion || '', ud: p.Ud || '',
      pu: p.PrecioVenta || 0,
      coste_ud: (costeSub[k] !== undefined ? costeSub[k] : 0),  // ASUNCIÓN: confirmar fuente
      te: c.te, ts: c.ts, ti: c.ti, to: c.to,
      _grupo: c._grupo, _tipo_grupo: c._tipo_grupo
    });
  });

  // 3) Materiales
  (DATOS.materiales || []).forEach(function (m) {
    var obra = OBRAS[(m.Proyecto || '').trim()];
    if (!obra) return;
    obra.mat.push({
      cod: m.Codigo || '', desc: m.Descripcion || '', ud: m.Ud || '',
      pu: m.PrecioUnitario || 0, coste_ud: m.CosteUnitario || 0,
      te: '', ts: '', ti: '', to: '', _grupo: m.Grupo || '', _tipo_grupo: 'SUM'
    });
  });

  // 4) Subcontratas
  (DATOS.subcontratas || []).forEach(function (s) {
    var obra = OBRAS[(s.Proyecto || '').trim()];
    if (!obra) return;
    obra.sub.push({
      task: s.TaskBC, cod: s.CodigoPartida || '', desc: s.Descripcion || '',
      ud: s.Ud || '', coste_estudio: s.CosteEstudio || 0, tarifa: s.Tarifa || 0,
      empresa: s.Empresa || '', contrato: s.Contrato || ''
    });
  });

  // 5) MO Externa / Alquileres / Indirectos / Planificación (tal cual, por obra)
  (DATOS.moExterna || []).forEach(function (x) {
    var o = OBRAS[(x.Proyecto || '').trim()]; if (o) o.mo_externa.push(x);
  });
  (DATOS.alquileres || []).forEach(function (x) {
    var o = OBRAS[(x.Proyecto || '').trim()]; if (o) o.alquileres.push(x);
  });
  (DATOS.indirectos || []).forEach(function (x) {
    var o = OBRAS[(x.Proyecto || '').trim()]; if (o) o.indirectos.push(x);
  });
  (DATOS.planificacion || []).forEach(function (x) {
    var o = OBRAS[(x.Proyecto || '').trim()]; if (o) o.planificacion.push(x);
  });

  // 6) Personal CORPORATIVO -> se copia a todas las obras (no lleva Proyecto)
  var KM_DEF = (opts.kmPorDefecto !== undefined ? opts.kmPorDefecto : 0.26);
  var personalCorp = (DATOS.personal || []).map(function (p) {
    return {
      num: p.Num != null ? String(p.Num) : '',
      nombre: p.Nombre || '', cat: p.Categoria || '', tipo: p.Tipo || '',
      ch: p.CosteHora || 0, dc: p.DietaCompleta || 0, md: p.MediaDieta || 0,
      km: KM_DEF
    };
  });
  Object.keys(OBRAS).forEach(function (cod) {
    OBRAS[cod].personal = personalCorp.map(function (x) { return Object.assign({}, x); });
  });

  // 7) Maquinaria CORPORATIVA -> referencia compartida
  var maquinariaCorp = (DATOS.maquinaria || []).map(function (m) {
    return {
      cod: m.Codigo || '', desc: m.Descripcion || '', categoria: m.Categoria || '',
      coste_hora: m.CosteHora || 0, coste_dia: m.CosteDia || 0
    };
  });

  return { OBRAS: OBRAS, MAQUINARIA: maquinariaCorp, _pendiente: ['partidas_cert', 'partidas_ppto'] };
}

global.TransformarTelice = { transformar: transformar, clasificaPartida: clasificaPartida };

})(window);
