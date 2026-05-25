'use strict';
/**
 * wipe-datos-prueba.js
 * Borra TODOS los items de Requerimientos, OrdenesCompra y Remisiones
 * para dejar la app en ceros antes del uso productivo.
 *
 * NO toca: Proveedores, Insumos, Proyectos, ConfiguracionApp.
 *
 * Uso:
 *   node src/scripts/wipe-datos-prueba.js                  → dry-run (solo cuenta)
 *   node src/scripts/wipe-datos-prueba.js --confirm        → ejecuta el borrado
 */

require('dotenv').config();
const fs   = require('fs');
const path = require('path');
const g    = require('../graphStorage');

const CONFIRM = process.argv.includes('--confirm');

const LISTAS_A_BORRAR = ['Requerimientos', 'OrdenesCompra', 'Remisiones'];

async function borrarTodoLista(siteId, nombre) {
  const lst = await g.getListByName(siteId, nombre);
  if (!lst) { console.log(`  ⚠ Lista "${nombre}" no existe, se omite.`); return 0; }
  const items = await g.getListItems(siteId, lst.id);
  console.log(`  → "${nombre}": ${items.length} items`);
  if (!CONFIRM) return items.length;
  let borrados = 0;
  for (const it of items) {
    try {
      await g.deleteListItem(siteId, lst.id, it.id);
      borrados++;
      if (borrados % 10 === 0) process.stdout.write(`    ${borrados}/${items.length}\r`);
    } catch (e) { console.warn(`    ✗ id ${it.id}: ${e.message.split('\n')[0]}`); }
  }
  console.log(`    ✓ ${borrados}/${items.length} eliminados.`);
  return borrados;
}

function resetearContadorLegacy() {
  const ruta = process.env.PATH_CONTADOR_OC || path.join(__dirname, '../../data/contador_oc.json');
  try {
    if (fs.existsSync(ruta)) {
      fs.writeFileSync(ruta, JSON.stringify({ ultimo: 0, actualizado: new Date().toISOString() }, null, 2));
      console.log(`  ✓ Contador legacy reseteado: ${ruta}`);
    } else {
      console.log(`  · No existe contador legacy en ${ruta}`);
    }
  } catch (e) { console.warn(`  ⚠ No se pudo resetear contador legacy: ${e.message}`); }
}

function limpiarTempLocal() {
  const dir = path.join(__dirname, '../../temp/cotizaciones');
  if (!fs.existsSync(dir)) return;
  const archivos = fs.readdirSync(dir);
  if (!CONFIRM) { console.log(`  → temp/cotizaciones: ${archivos.length} archivos`); return; }
  let n = 0;
  for (const a of archivos) {
    try { fs.unlinkSync(path.join(dir, a)); n++; } catch {}
  }
  console.log(`  ✓ ${n} archivos temporales eliminados.`);
}

async function main() {
  const host = process.env.SHAREPOINT_HOSTNAME;
  const sp   = process.env.SHAREPOINT_SITE_PATH;
  const site = await g.getSite(host, sp);

  console.log(CONFIRM ? '🔥 MODO EJECUCIÓN (--confirm)' : '👁  DRY-RUN (usa --confirm para borrar)');
  console.log('─'.repeat(60));

  let total = 0;
  for (const nombre of LISTAS_A_BORRAR) {
    total += await borrarTodoLista(site.id, nombre);
  }

  console.log('\n─ Archivos temporales locales ─');
  limpiarTempLocal();

  if (CONFIRM) {
    console.log('\n─ Contador OC ─');
    resetearContadorLegacy();
    console.log('  · El contador SharePoint (max+1 de OCs aprobadas) vuelve a 1 automáticamente al no quedar OCs.');
  }

  console.log('\n' + '─'.repeat(60));
  console.log(CONFIRM
    ? `✅ Wipe completado. Total items eliminados: ${total}`
    : `📋 Se eliminarían ${total} items. Ejecuta con --confirm para proceder.`);
  console.log('\nNota: los archivos PDF/XLSX subidos a OneDrive/SharePoint no se eliminan automáticamente.');
  console.log('      Si tenías adjuntos, bórralos manualmente desde la biblioteca de documentos del sitio.');
}

main().catch(err => { console.error('❌', err.message); process.exit(1); });
