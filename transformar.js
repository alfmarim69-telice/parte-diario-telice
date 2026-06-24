/* ============================================================
   transformar.js · Parte Diario Telice
   DATOS (11 maestros SharePoint) -> OBRAS (estructura de la app)
   ------------------------------------------------------------
   OPCIÓN A: SharePoint guarda los maestros crudos; la app deriva
   certificación, presupuesto de control y agrupaciones.

   MODELO DE PARTIDAS (estructura definitiva, validada contra
   23-039_Maestro_Obra.xlsx y 025-388_Maestro_Obra_DEF.xlsx):
   - M_Partidas = partidas de PRODUCCIÓN. Una fila = una unidad
     de venta. Cada partida se descompone en hasta 4 dimensiones
     de control (Ingeniería, Suministro, Ejecución, Otros), cada
     una con su CÓDIGO de control (Ctrl*) y sus precios unitarios
     propios:
       · PuVenta{Ing,Sum,Eje,Otr} = precio de venta unitario de esa fase
       · CosteUd{Ing,Sum,Eje,Otr} = coste unitario de esa fase
   - La suma de los PuVenta de las 4 dimensiones = PrecioVenta de
     la partida (precio de proyecto). Igual para el coste.
   - Totales por dimensión = PU * Medición (la app los calcula).
   - k de paso: ya NO es dato; se deriva del coste por dimensión
     (coste_total_dim / coste_total) para usos informativos.

   SALIDAS:
   - partidas      : operativa del parte diario (una por unidad)
   - partidas_cert : certificación (precio proyecto x medición), por capítulo
   - partidas_ppto : control económico, agrupado por código de control
                     (equivale a la pestaña Partidas_Control del maestro)

   Requiere SPTelice.leerTodo() (sp-integracion.js).
   ============================================================ */
(function (global) {
'use strict';

function r2(x) { return Math.round((x || 0) * 100) / 100; }
function num(x) { return (typeof x === 'number' && !isNaN(x)) ? x : 0; }

// Índice (proyecto|codigo) -> CosteEstudio, desde subcontratas.
function indiceCosteSub(subcontratas) {
  var idx = {};
  subcontratas.forEach(function (s) {
    var k = (s.Proyecto || '') + '|' + (s.CodigoPartida || '');
    if (idx[k] === undefined) idx[k] = s.CosteEstudio;
  });
  return idx;
}

// Las 4 dimensiones de una partida: código de control, PU venta y coste ud.
// Solo se devuelven las dimensiones que tienen código de control asignado.
function dimensiones(p) {
  var defs = [
    { cod: p.CtrlIngenieria, pu: p.PuVentaIng, cu: p.CosteUdIng, tipo: 'ING'   },
    { cod: p.CtrlSuministro, pu: p.PuVentaSum, cu: p.CosteUdSum, tipo: 'SUM'   },
    { cod: p.CtrlEjecucion,  pu: p.PuVentaEje, cu: p.CosteUdEje, tipo: 'EJEC'  },
    { cod: p.CtrlOtros,      pu: p.PuVentaOtr, cu: p.CosteUdOtr, tipo: 'OTROS' }
  ];
  return defs.filter(function (d) { return d.cod && String(d.cod).trim(); })
             .map(function (d) {
               return { cod: String(d.cod).trim(), pu: num(d.pu), cu: num(d.cu), tipo: d.tipo };
             });
}

// Clasificación dominante de una partida (grupo/tipo) para la operativa.
// Prioridad: ING > OTROS > SUM > EJEC.
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

  // Acumulador de control por obra: { obraCod: { codControl: {desc,venta,coste,tipo} } }
  var ctrlAcc = {};

  // 2) Partidas de producción -> operativa + certificación + acumulación de control
  (DATOS.partidas || []).forEach(function (p) {
    var obraCod = (p.Proyecto || '').trim();
    var obra = OBRAS[obraCod];
    if (!obra) return;

    var c   = clasificaPartida(p);
    var med = num(p.Medicion);
    var pu  = num(p.PrecioVenta);
    // coste_ud de la partida: CosteUd de M_Partidas; si no, subcontratas; si no, 0
    var k   = obraCod + '|' + (p.Codigo || '');
    var costeUd = (p.CosteUd !== undefined && p.CosteUd !== null && p.CosteUd !== '')
                  ? num(p.CosteUd)
                  : (costeSub[k] !== undefined ? num(costeSub[k]) : 0);

    var dims = dimensiones(p);

    // 2a) Operativa: la partida tal cual para el parte diario
    obra.partidas.push({
      cod: p.Codigo || '', desc: p.Descripcion || '', ud: p.Ud || '',
      pu: pu,
      coste_ud: costeUd,
      te: c.te, ts: c.ts, ti: c.ti, to: c.to,
      _grupo: c._grupo, _tipo_grupo: c._tipo_grupo
    });

    // 2b) Certificación: precio proyecto x medición, agrupable por Capítulo
    obra.partidas_cert.push({
      cod: p.Codigo || '', desc: p.Descripcion || '', ud: p.Ud || '',
      med: med,
      pu: pu,
      venta_ppto: r2(pu * med),
      coste_ud: costeUd,
      cap: p.Capitulo || '',
      ti: c.ti, ts: c.ts, te: c.te, to: c.to
    });

    // 2c) Control: cada dimensión aporta su (PU x med) y (coste_ud x med)
    //      al código de control correspondiente. Sin reparto por k: los
    //      precios unitarios por dimensión ya vienen desglosados.
    if (!ctrlAcc[obraCod]) ctrlAcc[obraCod] = {};
    var acc = ctrlAcc[obraCod];
    dims.forEach(function (d) {
      if (!acc[d.cod]) {
        acc[d.cod] = { cod: d.cod, desc: (p.Descripcion || ''), ud: 'PA',
                       venta: 0, coste: 0, tipo: d.tipo };
      }
      acc[d.cod].venta += d.pu * med;
      acc[d.cod].coste += d.cu * med;
    });
  });

  // 3) partidas_ppto = control económico agrupado (equivale a Partidas_Control)
  Object.keys(ctrlAcc).forEach(function (obraCod) {
    var obra = OBRAS[obraCod];
    if (!obra) return;
    var acc = ctrlAcc[obraCod];
    var codigos = Object.keys(acc).sort(function (a, b) {
      return a.localeCompare(b, undefined, { numeric: true });
    });
    codigos.forEach(function (cod) {
      var x = acc[cod];
      obra.partidas_ppto.push({
        cod: x.cod, desc: x.desc, ud: x.ud,
        med: 1,
        venta_ppto: r2(x.venta),
        coste_total_ppto: r2(x.coste),
        _tipo_grupo: x.tipo
      });
    });
  });

  // 4) M_PartidasControl en SP: SOLO sobrescribe la descripción "oficial"
  //    del código de control. El importe siempre se deriva de producción.
  (DATOS.partidasControl || []).forEach(function (pc) {
    var obra = OBRAS[(pc.Proyecto || '').trim()];
    if (!obra) return;
    var codigo = String(pc.Codigo || '').trim();
    var fila = obra.partidas_ppto.filter(function (r) { return r.cod === codigo; })[0];
    if (fila && pc.Descripcion) fila.desc = pc.Descripcion;
  });

  // 5) Materiales
  (DATOS.materiales || []).forEach(function (m) {
    var obra = OBRAS[(m.Proyecto || '').trim()];
    if (!obra) return;
    obra.mat.push({
      cod: m.Codigo || '', desc: m.Descripcion || '', ud: m.Ud || '',
      pu: num(m.PrecioUnitario), coste_ud: num(m.CosteUnitario),
      te: '', ts: '', ti: '', to: '', _grupo: m.Grupo || '', _tipo_grupo: 'SUM'
    });
  });

  // 6) Subcontratas
  (DATOS.subcontratas || []).forEach(function (s) {
    var obra = OBRAS[(s.Proyecto || '').trim()];
    if (!obra) return;
    obra.sub.push({
      task: s.TaskBC, cod: s.CodigoPartida || '', desc: s.Descripcion || '',
      ud: s.Ud || '', coste_estudio: num(s.CosteEstudio), tarifa: num(s.Tarifa),
      empresa: s.Empresa || '', contrato: s.Contrato || ''
    });
  });

  // 7) MO Externa / Alquileres / Indirectos / Planificación (tal cual, por obra)
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

  // 8) Personal CORPORATIVO -> se copia a todas las obras (no lleva Proyecto)
  var KM_DEF = (opts.kmPorDefecto !== undefined ? opts.kmPorDefecto : 0.26);
  var personalCorp = (DATOS.personal || []).map(function (p) {
    return {
      num: p.Num != null ? String(p.Num) : '',
      nombre: p.Nombre || '', cat: p.Categoria || '', tipo: p.Tipo || '',
      ch: num(p.CosteHora), dc: num(p.DietaCompleta), md: num(p.MediaDieta),
      km: KM_DEF
    };
  });
  Object.keys(OBRAS).forEach(function (cod) {
    OBRAS[cod].personal = personalCorp.map(function (x) { return Object.assign({}, x); });
  });

  // 9) Maquinaria CORPORATIVA -> referencia compartida
  var maquinariaCorp = (DATOS.maquinaria || []).map(function (m) {
    return {
      cod: m.Codigo || '', desc: m.Descripcion || '', categoria: m.Categoria || '',
      coste_hora: num(m.CosteHora), coste_dia: num(m.CosteDia)
    };
  });

  return { OBRAS: OBRAS, MAQUINARIA: maquinariaCorp };
}

global.TransformarTelice = {
  transformar: transformar,
  clasificaPartida: clasificaPartida,
  dimensiones: dimensiones
};
}(typeof window !== 'undefined' ? window : this));
