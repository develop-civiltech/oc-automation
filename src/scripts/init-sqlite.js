'use strict';
/**
 * init-sqlite.js — Migración inicial SharePoint → SQLite
 *
 * Ejecutar UNA VEZ para poblar la base de datos local:
 *   node src/scripts/init-sqlite.js
 *
 * En ejecuciones posteriores el syncService mantiene SQLite actualizado
 * automáticamente en segundo plano.
 */

require('dotenv').config({ path: require('path').join(__dirname, '../../.env') });

const g    = require('../graphStorage');
const db   = require('../db');
const sync = require('../syncService');

async function ctxSharePoint() {
  const host     = process.env.SHAREPOINT_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !sitePath) throw new Error('SHAREPOINT_HOSTNAME / SHAREPOINT_SITE_PATH no configurados en .env');
  const site = await g.getSite(host, sitePath);
  const ctx  = { siteId: site.id };
  for (const nombre of ['Requerimientos','OrdenesCompra','Insumos','Proveedores','Remisiones','Proyectos','OrdenesServicio','HistorialPrecios']) {
    try {
      const lst = await g.getListByName(site.id, nombre);
      if (lst) ctx[nombre] = lst.id;
    } catch {}
  }
  return ctx;
}

async function main() {
  console.log('\n── Inicializando base de datos SQLite ──────────────────────────────────');
  console.log(`   Archivo: ${process.env.SQLITE_PATH || './data/local.db'}\n`);

  if (!db.isReady()) {
    console.error('✗ No se pudo crear el archivo SQLite. Verifica permisos y la ruta SQLITE_PATH.');
    process.exit(1);
  }

  console.log('Conectando a SharePoint…');
  let ctx;
  try {
    ctx = await ctxSharePoint();
    console.log(`✓ Sitio: ${ctx.siteId}`);
  } catch (e) {
    console.error('✗ No se pudo conectar a SharePoint:', e.message);
    console.error('  Verifica SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, TENANT_ID, CLIENT_ID, CLIENT_SECRET en .env');
    process.exit(1);
  }

  console.log('\nDescargando listas…');
  const t0 = Date.now();
  await sync.init(() => Promise.resolve(ctx));

  // Esperar a que el primer sync termine (init lo dispara async)
  let intentos = 0;
  while (sync.isSyncing() && intentos++ < 60) {
    await new Promise(r => setTimeout(r, 1000));
    process.stdout.write('.');
  }
  process.stdout.write('\n');
  sync.stop();

  const duracion = ((Date.now() - t0) / 1000).toFixed(1);
  const conteos  = db.counts();

  console.log('\n── Resultado ───────────────────────────────────────────────────────────');
  const filas = [
    ['Lista SharePoint',    'Tabla SQLite',         'Registros'],
    ['HistorialPrecios',    'historial_precios',    conteos.historial_precios],
    ['Proveedores',         'proveedores',          conteos.proveedores],
    ['Insumos',             'insumos',              conteos.insumos],
    ['Proyectos',           'proyectos',            conteos.proyectos],
    ['Requerimientos',      'requerimientos',       conteos.requerimientos],
    ['OrdenesCompra',       'ordenes_compra',       conteos.ordenes_compra],
    ['OrdenesServicio',     'ordenes_servicio',     conteos.ordenes_servicio],
    ['Remisiones',          'remisiones',           conteos.remisiones],
  ];
  const w = [22, 22, 10];
  console.log(filas.map(r => r.map((c, i) => String(c).padEnd(w[i])).join(' ')).join('\n'));
  const total = Object.values(conteos).reduce((s, n) => s + n, 0);
  console.log(`\n   Total: ${total} registros en ${duracion} s`);

  const estados = db.getAllSyncState();
  if (estados.length) {
    console.log('\n── Estado de sincronización ────────────────────────────────────────────');
    for (const s of estados) {
      console.log(`   ${s.lista.padEnd(20)} ${s.item_count} items — ${s.last_sync}`);
    }
  }

  console.log('\n✓ SQLite listo. El servidor usará esta base de datos en lugar de llamar');
  console.log('  directamente a SharePoint en cada request.\n');
  process.exit(0);
}

main().catch(e => {
  console.error('\n✗ Error inesperado:', e.message);
  process.exit(1);
});
