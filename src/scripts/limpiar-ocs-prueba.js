'use strict';
/**
 * limpiar-ocs-prueba.js
 * Elimina de SharePoint y SQLite las OCs con estado='anulada' y sin número de consecutivo.
 * Estas OCs fueron generadas durante pruebas y nunca aprobadas.
 *
 * Uso:
 *   node src/scripts/limpiar-ocs-prueba.js              <- dry-run (solo lista, sin cambios)
 *   node src/scripts/limpiar-ocs-prueba.js --confirm    <- ejecuta el borrado
 */

require('dotenv').config();
const g       = require('../graphStorage');
const localDb = require('../db');

const CONFIRM = process.argv.includes('--confirm');

async function main() {
  console.log('Conectando a SharePoint...');
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const lst  = await g.getListByName(site.id, 'OrdenesCompra');
  if (!lst) { console.error('Lista OrdenesCompra no encontrada'); process.exit(1); }

  console.log('Leyendo OCs desde SharePoint...');
  const items = await g.getListItems(site.id, lst.id);

  const candidatas = items.filter(it => {
    const f = it.fields || {};
    return f.estado === 'anulada' && (!f.numeroOC || String(f.numeroOC).trim() === '');
  });

  console.log(`\nOCs anuladas sin número de consecutivo: ${candidatas.length}`);
  candidatas.forEach(it => {
    const f = it.fields || {};
    console.log(`  • id=${it.id} | proyecto=${f.proyecto || '—'} | proveedor=${f.proveedorNombre || '—'} | fecha=${(f.fechaCreacion || '').slice(0, 10) || '—'}`);
  });

  if (!candidatas.length) {
    console.log('\nNo hay registros para eliminar.');
    return;
  }

  if (!CONFIRM) {
    console.log('\nDry-run: ningún registro fue eliminado. Usa --confirm para ejecutar el borrado.');
    return;
  }

  console.log('\nEliminando en paralelo...');
  let ok = 0, err = 0;
  await Promise.all(candidatas.map(async it => {
    try {
      await g.deleteListItem(site.id, lst.id, it.id);
      localDb.db().prepare('DELETE FROM ordenes_compra WHERE sp_id = ?').run(String(it.id));
      ok++;
      console.log(`  ✓ ${it.id}`);
    } catch (e) {
      err++;
      console.warn(`  ✗ ${it.id}: ${e.message.split('\n')[0]}`);
    }
  }));

  console.log(`\nResultado: ${ok} eliminadas, ${err} errores.`);
}

main().catch(e => { console.error('Error fatal:', e.message); process.exit(1); });
