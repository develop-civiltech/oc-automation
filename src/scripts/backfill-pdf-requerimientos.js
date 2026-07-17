'use strict';
/**
 * backfill-pdf-requerimientos.js
 * Genera y sube a SharePoint (carpeta RequerimientosPDF) el PDF de los
 * requerimientos que ya existen en la lista pero que quedaron sin adjuntoUrl
 * (creados antes de que la generación automática de PDF existiera).
 *
 * Uso:
 *   node src/scripts/backfill-pdf-requerimientos.js            → dry-run (solo lista)
 *   node src/scripts/backfill-pdf-requerimientos.js --confirm  → ejecuta
 *   node src/scripts/backfill-pdf-requerimientos.js --confirm --todos
 *       → además de los que no tienen PDF, regenera los que ya tienen uno
 */

require('dotenv').config();
const requerimientos = require('../requerimientos');

const CONFIRM = process.argv.includes('--confirm');
const TODOS   = process.argv.includes('--todos');

async function main() {
  const items = await requerimientos.listar();
  const pendientes = TODOS ? items : items.filter(it => !it.adjuntoUrl);

  console.log(`Requerimientos totales: ${items.length}`);
  console.log(`Pendientes de PDF${TODOS ? ' (--todos: se regeneran todos)' : ''}: ${pendientes.length}`);

  if (!CONFIRM) {
    pendientes.forEach(it => console.log(`  - id ${it.id} · ${it.consecutivoSistema || it.consecutivo || '(sin consecutivo)'} · ${it.proyecto || '(sin proyecto)'}`));
    console.log('\nDry-run. Ejecuta con --confirm para generar y subir los PDFs.');
    return;
  }

  let ok = 0, fallidos = 0;
  for (const it of pendientes) {
    try {
      await requerimientos.regenerarPdf(it.id);
      ok++;
      console.log(`  ✓ id ${it.id} (${ok + fallidos}/${pendientes.length})`);
    } catch (e) {
      fallidos++;
      console.error(`  ✗ id ${it.id}: ${e.message}`);
    }
  }
  console.log(`\nCompletado — ok: ${ok} | fallidos: ${fallidos}`);
}

main().catch(e => { console.error('Fallo crítico:', e); process.exit(1); });
