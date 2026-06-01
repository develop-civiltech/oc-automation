'use strict';
/**
 * syncService.js — Sincronización SharePoint ↔ SQLite local
 *
 * - Al iniciar el servidor: descarga las 8 listas a SQLite.
 * - Cada SYNC_INTERVAL_MIN minutos: resync incremental.
 * - Endpoint GET /sync: fuerza resync manual.
 *
 * Lecturas → siempre desde SQLite (sin latencia de red).
 * Escrituras → SQLite primero (respuesta inmediata) + SP en segundo plano.
 */

require('dotenv').config();

const g  = require('./graphStorage');
const db = require('./db');

const SYNC_INTERVAL_MS = (parseFloat(process.env.SYNC_INTERVAL_MIN) || 0.167) * 60 * 1000; // default ~10 s

let _syncTimer    = null;
let _syncing      = false;
let _lastSyncMs   = 0;
let _ctxFn        = null; // inyectado desde servidor-cotizaciones.js

// ── Registro de listas a sincronizar ─────────────────────────────────────────

const LISTAS = [
  {
    nombre:   'HistorialPrecios',
    tabla:    null,
    upsert:   (rows) => db.bulkUpsertHistorialAsync(rows.map(it => ({ id: it.id, ...(it.fields || {}) }))),
  },
  {
    nombre:   'Proveedores',
    tabla:    null,
    upsert:   (rows) => db.bulkUpsertProveedores(rows.map(it => ({ id: it.id, ...(it.fields || {}) }))),
  },
  {
    nombre:   'Insumos',
    tabla:    null,
    upsert:   (rows) => db.bulkUpsertInsumos(rows.map(it => ({ id: it.id, ...(it.fields || {}) }))),
  },
  {
    nombre:   'Proyectos',
    tabla:    null,
    upsert:   (rows) => db.bulkUpsertProyectos(rows.map(it => ({ id: it.id, ...(it.fields || {}) }))),
  },
  {
    nombre:   'Requerimientos',
    tabla:    'requerimientos',
    upsert:   (rows) => db.bulkUpsertDocs('requerimientos', rows),
  },
  {
    nombre:   'OrdenesCompra',
    tabla:    'ordenes_compra',
    upsert:   (rows) => db.bulkUpsertDocs('ordenes_compra', rows),
  },
  {
    nombre:   'OrdenesServicio',
    tabla:    'ordenes_servicio',
    upsert:   (rows) => db.bulkUpsertDocs('ordenes_servicio', rows),
  },
  {
    nombre:   'Remisiones',
    tabla:    'remisiones',
    upsert:   (rows) => db.bulkUpsertDocs('remisiones', rows),
  },
  {
    nombre:   'MovimientosInventario',
    tabla:    'movimientos_inventario',
    upsert:   (rows) => db.bulkUpsertDocs('movimientos_inventario', rows),
  },
  {
    nombre:   'UsuariosERP',
    tabla:    null,
    upsert:   (rows) => db.bulkUpsertUsuarios(rows.map(it => ({ id: it.id, ...(it.fields || {}) }))),
  },
];

// ── Sincronización de una lista ───────────────────────────────────────────────

async function syncLista(ctx, lista) {
  const listId = ctx[lista.nombre];
  if (!listId) return { nombre: lista.nombre, count: 0, skipped: true };

  const rows = await g.getListItems(ctx.siteId, listId);
  await lista.upsert(rows);
  db.setSyncState(lista.nombre, rows.length);
  return { nombre: lista.nombre, count: rows.length };
}

// ── Sincronización completa ───────────────────────────────────────────────────

async function syncAll() {
  if (_syncing) return { skipped: true, reason: 'sync ya en curso' };
  _syncing = true;
  const inicio = Date.now();
  const resultados = [];
  const errores    = [];

  try {
    const ctx = await _ctxFn();
    await Promise.allSettled(
      LISTAS.map(lista =>
        syncLista(ctx, lista)
          .then(r  => resultados.push(r))
          .catch(e => errores.push({ nombre: lista.nombre, error: e.message }))
      )
    );
    _lastSyncMs = Date.now();
  } catch (e) {
    errores.push({ nombre: 'ctxSharePoint', error: e.message });
  } finally {
    _syncing = false;
  }

  const duracion = Date.now() - inicio;
  const ok = errores.length === 0;
  const logMsg = `[syncService] ${ok ? '✓' : '⚠'} sync ${duracion} ms — ` +
    resultados.map(r => `${r.nombre}:${r.count ?? '?'}`).join(', ') +
    (errores.length ? ` | errores: ${errores.map(e => e.nombre).join(',')}` : '');
  console.log(logMsg);

  return { ok, duracion, resultados, errores };
}

// ── API pública ───────────────────────────────────────────────────────────────

/**
 * Debe llamarse una vez al iniciar el servidor.
 * ctxFn: función async que retorna el contexto de SharePoint (ctxSharePoint).
 */
async function init(ctxFn) {
  _ctxFn = ctxFn;

  // Primer sync al iniciar (no bloquea el arranque del servidor)
  syncAll().catch(e => console.warn('[syncService] Sync inicial falló:', e.message));

  // Sync periódico en segundo plano
  _syncTimer = setInterval(() => {
    syncAll().catch(e => console.warn('[syncService] Sync periódico falló:', e.message));
  }, SYNC_INTERVAL_MS);

  // setInterval no debe bloquear el proceso
  if (_syncTimer.unref) _syncTimer.unref();
}

function stop() {
  if (_syncTimer) { clearInterval(_syncTimer); _syncTimer = null; }
}

function lastSync() { return _lastSyncMs ? new Date(_lastSyncMs).toISOString() : null; }

function isSyncing() { return _syncing; }

module.exports = { init, stop, syncAll, lastSync, isSyncing };
