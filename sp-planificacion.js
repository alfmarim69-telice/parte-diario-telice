/**
 * sp-planificacion.js
 * Parte Diario Telice — Integración SharePoint para Planificación
 * ---------------------------------------------------------------
 * Gestiona lectura y escritura de la lista "Planificacion" en SharePoint.
 *
 * PREREQUISITO: sp-integracion.js debe estar cargado antes que este script,
 * ya que reutiliza getSpToken() y SP_CONFIG definidos en él.
 *
 * LISTA SHAREPOINT REQUERIDA: "Planificacion"
 * Columnas necesarias (crearlas si no existen):
 *   - Title       (texto, una línea)  → se usa como campo Clave: "PROYECTO|COD|FECHA"
 *   - Proyecto    (texto, una línea)
 *   - Fecha       (fecha)
 *   - Semana      (número)
 *   - CodProduccion (texto, una línea)
 *   - Descripcion (texto, varias líneas o una línea)
 *   - Ud          (texto, una línea)
 *   - CantidadPlanif (número)
 *
 * ESTRATEGIA DE ESCRITURA (upsert por quincena):
 *   1. Lee todos los items existentes de la lista para esa obra y rango de fechas.
 *   2. Borra los que ya no están en el nuevo plan (líneas borradas por el usuario).
 *   3. Para cada fila del nuevo plan:
 *      - Si ya existe (mismo Clave = Title) → PATCH para actualizar CantidadPlanif.
 *      - Si no existe → POST para crear.
 *
 * LECTURA (cargar plan guardado):
 *   Devuelve un objeto PLAN { codPartida: { fechaISO: cantidad } }
 *   equivalente al estado interno de planificacion-prototipo_2.html.
 */

// ─── Configuración ────────────────────────────────────────────────────────────
// SP_CONFIG y getSpToken() vienen de sp-integracion.js.
// Si se usa este script de forma independiente, descomenta y completa:
//
// const SP_CONFIG = {
//   siteId: 'SITE-ID-DE-TU-SP',  // obtenido de sp-integracion.js
//   hostname: 'telice.sharepoint.com',
//   sitePath: '/sites/ParteDiario'
// };
// async function getSpToken() { /* ver sp-integracion.js */ }

const LISTA_PLANIFICACION = 'Planificacion';

// ─── Helpers ─────────────────────────────────────────────────────────────────

function isoAFecha(isoStr) {
  // SharePoint espera fechas en formato ISO 8601 completo para campos de tipo Date
  return isoStr + 'T00:00:00Z';
}

function buildClave(proyecto, cod, fecha) {
  return `${proyecto}|${cod}|${fecha}`;
}

/**
 * Construye la URL base del endpoint REST de la lista en Graph.
 */
function listaUrl() {
  return `https://graph.microsoft.com/v1.0/sites/${SP_CONFIG.siteId}/lists/${LISTA_PLANIFICACION}/items`;
}

/**
 * Ejecuta una petición autenticada contra Graph.
 * @param {string} url
 * @param {object} opts - opciones fetch adicionales (method, body, headers extra)
 */
async function spFetch(url, opts = {}) {
  const token = await getSpToken();
  const res = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(opts.headers || {})
    }
  });
  if (!res.ok) {
    const err = await res.text().catch(() => res.statusText);
    throw new Error(`SP ${opts.method || 'GET'} ${url} → ${res.status}: ${err}`);
  }
  // DELETE devuelve 204 sin body
  if (res.status === 204) return null;
  return res.json();
}

// ─── LECTURA ─────────────────────────────────────────────────────────────────

/**
 * Carga el plan guardado en SharePoint para una obra y rango de fechas (quincena).
 * @param {string} proyecto - código de obra, ej. "23-039"
 * @param {string} fechaIni - ISO date string, primer día de la quincena
 * @param {string} fechaFin - ISO date string, último día de la quincena
 * @returns {Object} PLAN = { "COD_PARTIDA": { "2026-07-01": 120, ... }, ... }
 */
async function leerPlanificacionDeSP(proyecto, fechaIni, fechaFin) {
  // Filtramos por Proyecto y rango de fechas usando OData $filter
  const filter = encodeURIComponent(
    `fields/Proyecto eq '${proyecto}' and fields/Fecha ge '${isoAFecha(fechaIni)}' and fields/Fecha le '${isoAFecha(fechaFin)}'`
  );
  const select = encodeURIComponent('fields/Title,fields/CodProduccion,fields/Fecha,fields/CantidadPlanif');
  const url = `${listaUrl()}?$filter=${filter}&$select=${select}&$top=2000`;

  const data = await spFetch(url);
  const plan = {};

  (data.value || []).forEach(item => {
    const f = item.fields;
    const cod = f.CodProduccion;
    // La fecha viene como "2026-07-01T00:00:00Z", nos quedamos con los 10 primeros chars
    const fecha = (f.Fecha || '').slice(0, 10);
    const cant = parseFloat(f.CantidadPlanif) || 0;

    if (cod && fecha && cant > 0) {
      if (!plan[cod]) plan[cod] = {};
      plan[cod][fecha] = cant;
    }
  });

  return plan;
}

// ─── ESCRITURA ────────────────────────────────────────────────────────────────

/**
 * Guarda (upsert) las filas de planificación en SharePoint.
 * Borra líneas que el usuario eliminó, actualiza las modificadas, crea las nuevas.
 *
 * @param {Array} filas - array de objetos generado por guardar() en el HTML:
 *   [ { Clave, Proyecto, Fecha, Semana, CodProduccion, Descripcion, Ud, CantidadPlanif }, ... ]
 * @param {string} fechaIni - ISO date del primer día de la quincena (para borrar huérfanos)
 * @param {string} fechaFin - ISO date del último día de la quincena
 * @returns {{ creados: number, actualizados: number, borrados: number, errores: string[] }}
 */
async function guardarPlanificacionEnSP(filas, fechaIni, fechaFin) {
  const resultado = { creados: 0, actualizados: 0, borrados: 0, errores: [] };

  if (!filas || filas.length === 0 && !fechaIni) {
    resultado.errores.push('Sin datos para guardar');
    return resultado;
  }

  const proyecto = filas.length > 0 ? filas[0].Proyecto : null;

  // ── 1. Leer items existentes en SP para este proyecto y quincena ──────────
  let itemsExistentes = [];
  try {
    const filter = proyecto
      ? encodeURIComponent(`fields/Proyecto eq '${proyecto}' and fields/Fecha ge '${isoAFecha(fechaIni)}' and fields/Fecha le '${isoAFecha(fechaFin)}'`)
      : encodeURIComponent(`fields/Fecha ge '${isoAFecha(fechaIni)}' and fields/Fecha le '${isoAFecha(fechaFin)}'`);

    const url = `${listaUrl()}?$filter=${filter}&$select=id,fields/Title&$top=2000`;
    const data = await spFetch(url);
    itemsExistentes = data.value || [];
  } catch (e) {
    resultado.errores.push(`Error leyendo SP: ${e.message}`);
    return resultado;
  }

  // Mapa: Clave (Title) → id de item en SP
  const mapaExistentes = {};
  itemsExistentes.forEach(item => {
    mapaExistentes[item.fields.Title] = item.id;
  });

  // Conjunto de claves nuevas
  const clavesNuevas = new Set(filas.map(f => f.Clave));

  // ── 2. Borrar items que ya no están en el plan ─────────────────────────────
  const aEliminar = itemsExistentes.filter(item => !clavesNuevas.has(item.fields.Title));
  for (const item of aEliminar) {
    try {
      await spFetch(`${listaUrl()}/${item.id}`, { method: 'DELETE' });
      resultado.borrados++;
    } catch (e) {
      resultado.errores.push(`Error borrando ${item.fields.Title}: ${e.message}`);
    }
  }

  // ── 3. Crear o actualizar cada fila del plan ──────────────────────────────
  for (const fila of filas) {
    const spFields = {
      Title: fila.Clave,
      Proyecto: fila.Proyecto,
      Fecha: isoAFecha(fila.Fecha),
      Semana: fila.Semana,
      CodProduccion: fila.CodProduccion,
      Descripcion: fila.Descripcion,
      Ud: fila.Ud,
      CantidadPlanif: fila.CantidadPlanif
    };

    try {
      if (mapaExistentes[fila.Clave]) {
        // PATCH — actualizar
        await spFetch(`${listaUrl()}/${mapaExistentes[fila.Clave]}`, {
          method: 'PATCH',
          body: JSON.stringify({ fields: spFields })
        });
        resultado.actualizados++;
      } else {
        // POST — crear
        await spFetch(listaUrl(), {
          method: 'POST',
          body: JSON.stringify({ fields: spFields })
        });
        resultado.creados++;
      }
    } catch (e) {
      resultado.errores.push(`Error guardando ${fila.Clave}: ${e.message}`);
    }
  }

  return resultado;
}

// ─── Exportar para uso desde el HTML ─────────────────────────────────────────
// Si se usa como módulo ES: export { leerPlanificacionDeSP, guardarPlanificacionEnSP };
// En HTML con <script src="sp-planificacion.js">, las funciones quedan globales.
