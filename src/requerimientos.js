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

const g  = require('./graphStorage');
const db = require('./db');

const _cache = {}; // { siteId, listId, proyectosListId }
async function ctx() {
  if (_cache.listId) return _cache;
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const list = await g.getListByName(site.id, 'Requerimientos');
  if (!list) throw new Error('List "Requerimientos" no existe. Ejecuta crear-listas.js');
  _cache.siteId = site.id;
  _cache.listId = list.id;
  // Lista Proyectos — para el contador de consecutivos en SP
  const proyList = await g.getListByName(site.id, 'Proyectos').catch(() => null);
  _cache.proyectosListId = proyList?.id || null;
  // Migraciones no destructivas
  try { await g.addListColumn(site.id, list.id, { name: 'consecutivoSistema', text: {} }); } catch {}
  if (_cache.proyectosListId) {
    try { await g.addListColumn(site.id, _cache.proyectosListId, { name: 'ultimoConsecutivoReq', number: {} }); } catch {}
  }
  return _cache;
}

// Obtiene el siguiente consecutivo desde SP Proyectos (fuente de verdad) con etag.
// Retorna null si el proyecto no existe en SP o si ocurre un error no recuperable.
async function getNextConsecutivoDesdeProyectosSP(siteId, proyectosListId, proyecto) {
  if (!proyectosListId || !proyecto) return null;
  try {
    const items = await g.getListItems(siteId, proyectosListId, {
      filter: `fields/nombre eq '${proyecto.replace(/'/g, "''")}'`,
      prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly',
      top: 1,
    });
    if (!items.length) return null;

    const item      = items[0];
    const etag      = item['@odata.etag'] || item.eTag;
    const actual    = Number(item.fields?.ultimoConsecutivoReq || 0);
    const siguiente = actual + 1;

    await g.updateListItem(siteId, proyectosListId, item.id,
      { ultimoConsecutivoReq: siguiente },
      etag ? { etag } : {}
    );

    // Sync local como cache (best-effort)
    try { db.setConsecutivoProyecto(proyecto, siguiente); } catch {}

    return String(siguiente).padStart(4, '0');
  } catch (e) {
    if (e.status === 412) {
      // Conflicto de concurrencia — otro usuario llegó primero, reintentar
      return getNextConsecutivoDesdeProyectosSP(siteId, proyectosListId, proyecto);
    }
    return null; // Otro error → el llamador usa fallback SQLite
  }
}

// ── Mapeo del resultado de procesarCorreo → fields del item ──────────────────

function mapearFields(resultado, { messageId = '', adjuntoUrl = '', consecutivoSistema = '' } = {}) {
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
    consecutivo:         s.consecutivo || '',
    consecutivoSistema:  consecutivoSistema,
    proyecto:            s.proyecto || '',
    fechaSolicitud:      s.fechaSolicitud ? fechaISO(s.fechaSolicitud) : new Date().toISOString(),
    solicitante:         s.responsable || '',
    estado:              'pendiente',
    origenCorreoId:      messageId,
    adjuntoUrl:          adjuntoUrl,
    itemsJson:           JSON.stringify(items),
    notas:               (resultado.alertasGlobales || []).join(' | '),
    ocsGeneradas:        '',
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
  const { siteId, listId, proyectosListId } = await ctx();

  // Deduplicar por messageId — solo para correos reales (no cargas manuales).
  // Los messageId manuales tienen formato "manual:..." y son siempre únicos
  // (incluyen timestamp), por lo que no requieren consulta a SharePoint.
  // Para correos reales, origenCorreoId no está indexado en SP → necesita Prefer header.
  if (meta.messageId && !meta.messageId.startsWith('manual:')) {
    const existentes = await g.getListItems(siteId, listId, {
      filter: `fields/origenCorreoId eq '${meta.messageId.replace(/'/g, "''")}'`,
      prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly',
    });
    if (existentes.length > 0) return { item: existentes[0], duplicado: true };
  }

  // Consecutivo de sistema: fuente de verdad en SP Proyectos, fallback a SQLite local
  const proyecto = (resultado.solicitud || {}).proyecto || '';
  const consecutivoSistema = proyecto
    ? (await getNextConsecutivoDesdeProyectosSP(siteId, proyectosListId, proyecto)
       || db.getNextConsecutivoProyecto(proyecto))
    : '';

  const fields = mapearFields(resultado, { ...meta, consecutivoSistema });

  const item = await g.addListItem(siteId, listId, fields);
  return { item, duplicado: false, consecutivoSistema };
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
