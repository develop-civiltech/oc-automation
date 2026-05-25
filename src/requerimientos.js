'use strict';
/**
 * requerimientos.js
 * CRUD de la List "Requerimientos" en SharePoint. En el nuevo flujo, los
 * correos entrantes ya no generan OC automáticamente; crean un Requerimiento
 * en estado 'pendiente' para que el usuario lo gestione desde la consola.
 *
 * API:
 *   crearDesdeCorreo(resultado, { messageId, adjuntoUrl })  → item creado
 *   listar()                                                 → items activos
 *   actualizar(itemId, cambios, etag?)                       → PATCH
 *   marcarGestionado(itemId, ocsGeneradas)                   → estado gestionado
 *   bloquear(itemId, usuario, minutos = 15)                  → soft-lock
 */

const g = require('./graphStorage');

const _cache = {}; // { siteId, listId }
async function ctx() {
  if (_cache.listId) return _cache;
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const list = await g.getListByName(site.id, 'Requerimientos');
  if (!list) throw new Error('List "Requerimientos" no existe. Ejecuta crear-listas.js');
  _cache.siteId = site.id;
  _cache.listId = list.id;
  return _cache;
}

// ── Mapeo del resultado de procesarCorreo → fields del item ──────────────────

function mapearFields(resultado, { messageId = '', adjuntoUrl = '' } = {}) {
  const s = resultado.solicitud || {};
  const items = (resultado.items || []).map(it => ({
    insumo:    it.insumo,
    cantidad:  it.cantidad,
    unidad:    it.unidad,
    necesidad: it.necesidad || '',
    posibleProveedor: it.posibleProveedor || '',
    consulta: it.consulta ? {
      encontrado: !!it.consulta.encontrado,
      precio:     it.consulta.precio || 0,
      proveedor:  it.consulta.proveedor || null,
      alertas:    it.consulta.alertas || [],
      sinHistorial: !!it.consulta.sinHistorial,
    } : null,
  }));

  return {
    consecutivo:    s.consecutivo || '',
    proyecto:       s.proyecto || '',
    fechaSolicitud: s.fechaSolicitud ? fechaISO(s.fechaSolicitud) : new Date().toISOString(),
    solicitante:    s.responsable || '',
    estado:         'pendiente',
    origenCorreoId: messageId,
    adjuntoUrl:     adjuntoUrl,
    itemsJson:      JSON.stringify(items),
    notas:          (resultado.alertasGlobales || []).join(' | '),
    ocsGeneradas:   '',
  };
}

function fechaISO(f) {
  // Acepta "DD/MM/YYYY" o Date o ISO
  if (f instanceof Date) return f.toISOString();
  if (typeof f !== 'string') return new Date().toISOString();
  const m = f.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(`${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T00:00:00Z`).toISOString();
  const d = new Date(f);
  return isNaN(d) ? new Date().toISOString() : d.toISOString();
}

// ── API pública ───────────────────────────────────────────────────────────────

async function crearDesdeCorreo(resultado, meta = {}) {
  const { siteId, listId } = await ctx();
  const fields = mapearFields(resultado, meta);

  // Deduplicar: si ya existe un Requerimiento con el mismo consecutivo + proyecto,
  // no crear otro (correos reenviados, retries, etc.).
  if (fields.consecutivo) {
    const existentes = await g.getListItems(siteId, listId, {
      filter: `fields/consecutivo eq '${fields.consecutivo.replace(/'/g,"''")}'`,
    });
    const dup = existentes.find(e => (e.fields || {}).proyecto === fields.proyecto);
    if (dup) return { item: dup, duplicado: true };
  }

  const item = await g.addListItem(siteId, listId, fields);
  return { item, duplicado: false };
}

async function listar(filter) {
  const { siteId, listId } = await ctx();
  const opts = filter ? { filter } : {};
  const items = await g.getListItems(siteId, listId, opts);
  return items.map(it => ({ id: it.id, ...(it.fields || {}) }));
}

async function actualizar(itemId, cambios, etag) {
  const { siteId, listId } = await ctx();
  return g.updateListItem(siteId, listId, itemId, cambios, etag ? { etag } : {});
}

async function marcarGestionado(itemId, ocsGeneradas = []) {
  return actualizar(itemId, {
    estado: 'gestionado',
    ocsGeneradas: Array.isArray(ocsGeneradas) ? ocsGeneradas.join(', ') : String(ocsGeneradas),
  });
}

async function bloquear(itemId, usuario, minutos = 15) {
  const hasta = new Date(Date.now() + minutos * 60 * 1000).toISOString();
  return actualizar(itemId, { bloqueadoPor: usuario, bloqueadoHasta: hasta });
}

async function liberar(itemId) {
  return actualizar(itemId, { bloqueadoPor: '', bloqueadoHasta: null });
}

module.exports = {
  crearDesdeCorreo, listar, actualizar,
  marcarGestionado, bloquear, liberar,
  mapearFields,
};
