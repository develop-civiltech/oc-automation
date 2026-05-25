'use strict';
/**
 * cargar-insumos.js
 * Sube el JSON validado (insumos_validados.json) a la List "Insumos" en SharePoint.
 *
 * Uso:
 *   node src/scripts/cargar-insumos.js [ruta_json]
 *   node src/scripts/cargar-insumos.js --limpiar              (borra items existentes antes)
 *   node src/scripts/cargar-insumos.js --dry-run
 *
 * Si no se pasa ruta, por defecto lee data/insumos_validados.json
 * (o data/insumos_sugeridos.json si el validado no existe — útil si se salta la UI).
 */

require('dotenv').config();
const fs   = require('fs');
const path = require('path');
const g    = require('../graphStorage');

const DRY_RUN  = process.argv.includes('--dry-run');
const LIMPIAR  = process.argv.includes('--limpiar');
const RUTA_ARG = process.argv.find(a => a.endsWith('.json'));
const RUTA_DEFAULT_VAL = path.join(__dirname, '../../data/insumos_validados.json');
const RUTA_DEFAULT_SUG = path.join(__dirname, '../../data/insumos_sugeridos.json');

function resolverRuta() {
  if (RUTA_ARG) return path.resolve(RUTA_ARG);
  if (fs.existsSync(RUTA_DEFAULT_VAL)) return RUTA_DEFAULT_VAL;
  if (fs.existsSync(RUTA_DEFAULT_SUG)) {
    console.log(`⚠  No existe insumos_validados.json; usando insumos_sugeridos.json`);
    return RUTA_DEFAULT_SUG;
  }
  throw new Error('No se encontró JSON de insumos. Genera con extraer-insumos.js o valida en UI.');
}

function normalizarItem(it) {
  return {
    nombre:            (it.nombre || '').toString().trim().substring(0, 255),
    nombreNormalizado: (it.nombreNormalizado || '').toString().trim().substring(0, 255),
    categoria:         it.categoria === '_SIN_CATEGORIA' ? '' : (it.categoria || ''),
    subcategoria:      (it.subcategoria || '').toString().trim().substring(0, 100),
    unidadEstandar:    (it.unidadEstandar || '').toString().trim().substring(0, 50),
    sinonimos:         (it.sinonimos || '').toString(),
    activo:            true,
  };
}

async function insertarLote(siteId, listId, items) {
  const requests = items.map((it, idx) => ({
    id:     String(idx + 1),
    method: 'POST',
    url:    `/sites/${siteId}/lists/${listId}/items`,
    headers: { 'Content-Type': 'application/json' },
    body:   { fields: it },
  }));

  const resp = await g.post('/$batch', { requests });
  const errores = (resp.responses || []).filter(r => r.status >= 400);
  return { total: items.length, errores };
}

async function main() {
  const ruta = resolverRuta();
  const raw  = JSON.parse(fs.readFileSync(ruta, 'utf8'));
  console.log(`Leyendo ${ruta}: ${raw.length} insumos`);

  const items = raw
    .filter(it => it.nombre && it.nombre.trim())
    .map(normalizarItem);

  const sinCat = items.filter(i => !i.categoria).length;
  if (sinCat > 0) {
    console.log(`⚠  ${sinCat} insumos sin categoría (se cargarán igual con categoria='')`);
  }

  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const list = await g.getListByName(site.id, 'Insumos');
  if (!list) throw new Error('List "Insumos" no encontrada');
  console.log(`List: Insumos (${list.id})`);

  if (LIMPIAR) {
    console.log(`→ Eliminando items existentes (--limpiar)...`);
    const actuales = await g.getListItems(site.id, list.id);
    console.log(`  ${actuales.length} items existentes`);
    if (!DRY_RUN) {
      for (const it of actuales) {
        await g.deleteListItem(site.id, list.id, it.id);
      }
      console.log(`  ✓ Eliminados`);
    }
  }

  console.log(`\nCargando ${items.length} insumos...`);
  const LOTE = 20;
  let ok = 0, fail = 0;

  for (let i = 0; i < items.length; i += LOTE) {
    const lote = items.slice(i, i + LOTE);
    if (DRY_RUN) {
      console.log(`  [dry-run] Lote ${i/LOTE + 1}: ${lote.length} items`);
      continue;
    }
    const r = await insertarLote(site.id, list.id, lote);
    ok   += r.total - r.errores.length;
    fail += r.errores.length;
    console.log(`  Lote ${Math.floor(i/LOTE) + 1} [${ok + fail}/${items.length}]  errores: ${r.errores.length}`);
    if (r.errores.length) {
      for (const e of r.errores.slice(0, 3)) {
        console.log(`    ✗ ${e.status}: ${JSON.stringify(e.body).substring(0, 180)}`);
      }
    }
  }

  console.log(`\n✅ Carga completada. OK: ${ok}   Errores: ${fail}`);
}

main().catch(err => {
  console.error('❌ Error fatal:', err.message);
  if (err.response) console.error(JSON.stringify(err.response, null, 2));
  process.exit(1);
});
