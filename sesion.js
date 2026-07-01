/**
 * sesion.js — Parte Diario Telice
 * ─────────────────────────────────────────────────────────────────────────────
 * Gestión de sesión de usuario y control de acceso por rol.
 *
 * ROLES DISPONIBLES:
 *   'Jefe Equipo'            → puede crear y enviar partes
 *   'Jefe Obra'              → puede aprobar, devolver y reabrir partes
 *   'Jefe Departamento'      → puede ver bandeja de pendientes de su dpto.
 *   'Direccion'              → visibilidad total sobre todos los proyectos
 *   'Departamento Planificacion' → acceso a módulo de planificación
 *
 * USO EN index.html:
 *   1. Añadir <script src="sesion.js"></script> antes del cierre de </body>
 *   2. Llamar a Sesion.init() al arrancar la app
 *   3. Usar Sesion.get() para obtener el usuario activo
 *   4. Usar Sesion.puede(accion) para controlar botones
 */

const Sesion = (() => {

  const KEY = 'pd_usuario';

  // ─── Permisos por rol ─────────────────────────────────────────────────────
  const PERMISOS = {
    'Jefe Equipo': [
      'crear_parte',
      'editar_parte_propio',
      'enviar_parte',
    ],
    'Jefe Obra': [
      'crear_parte',
      'editar_parte_propio',
      'enviar_parte',
      'aprobar_parte',
      'devolver_parte',
      'reabrir_parte',
      'ver_pendientes',
      'ver_todos_partes',
    ],
    'Jefe Departamento': [
      'ver_pendientes',
      'ver_todos_partes',
      'aprobar_parte',
      'devolver_parte',
    ],
    'Dirección': [
      'crear_parte',
      'editar_parte_propio',
      'enviar_parte',
      'aprobar_parte',
      'devolver_parte',
      'reabrir_parte',
      'ver_pendientes',
      'ver_todos_partes',
      'ver_informes',
      'ver_planificacion',
    ],
    'Departamento Planificacion': [
      'ver_todos_partes',
      'ver_informes',
      'ver_planificacion',
      'editar_planificacion',
    ],
  };

  // ─── API pública ──────────────────────────────────────────────────────────

  /**
   * Inicializa la sesión. Si no hay sesión activa redirige a login.html.
   * Llamar al arrancar index.html.
   */
  function init() {
    const sesion = get();
    if (!sesion) {
      window.location.href = 'login.html';
      return null;
    }
    _aplicarUI(sesion);
    return sesion;
  }

  /** Devuelve el objeto de sesión activo o null. */
  function get() {
    const raw = sessionStorage.getItem(KEY);
    if (!raw) return null;
    try { return JSON.parse(raw); } catch { return null; }
  }

  /** Cierra la sesión y redirige a login. */
  function cerrar() {
    sessionStorage.removeItem(KEY);
    window.location.href = 'login.html';
  }

  /**
   * Comprueba si el usuario activo tiene permiso para una acción.
   * @param {string} accion - ej. 'aprobar_parte', 'ver_pendientes'
   * @returns {boolean}
   */
  function puede(accion) {
    const sesion = get();
    if (!sesion) return false;
    const permisos = PERMISOS[sesion.rol] || [];
    return permisos.includes(accion);
  }

  /** Devuelve el rol del usuario activo o null. */
  function rol() {
    const sesion = get();
    return sesion ? sesion.rol : null;
  }

  /** Devuelve el nombre del usuario activo o null. */
  function nombre() {
    const sesion = get();
    return sesion ? sesion.nombre : null;
  }

  // ─── Aplicar UI según rol ─────────────────────────────────────────────────

  /**
   * Muestra/oculta elementos del DOM según el rol del usuario.
   * Usa atributos data-rol-min y data-accion en los elementos HTML:
   *
   *   <button data-accion="aprobar_parte">Aprobar</button>
   *   → visible solo si el rol tiene permiso 'aprobar_parte'
   *
   *   <div data-rol="Direccion,Jefe Departamento">...</div>
   *   → visible solo para esos roles exactos
   */
  function _aplicarUI(sesion) {
    // Por acción
    document.querySelectorAll('[data-accion]').forEach(el => {
      const accion = el.dataset.accion;
      const permisos = PERMISOS[sesion.rol] || [];
      el.style.display = permisos.includes(accion) ? '' : 'none';
    });

    // Por rol explícito
    document.querySelectorAll('[data-rol]').forEach(el => {
      const roles = el.dataset.rol.split(',').map(r => r.trim());
      el.style.display = roles.includes(sesion.rol) ? '' : 'none';
    });

    // Mostrar nombre de usuario en elementos con id="usuarioActivo"
    const elNombre = document.getElementById('usuarioActivo');
    if (elNombre) {
      elNombre.textContent = `${sesion.nombre} · ${sesion.rol}`;
    }

    // Mostrar inicial en avatar con id="usuarioAvatar"
    const elAvatar = document.getElementById('usuarioAvatar');
    if (elAvatar) {
      elAvatar.textContent = sesion.nombre ? sesion.nombre.charAt(0).toUpperCase() : '?';
    }
  }

  return { init, get, cerrar, puede, rol, nombre };

})();
