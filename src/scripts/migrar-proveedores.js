'use strict';
/**
 * migrar-proveedores.js
 * Carga el CSV data/proveedores_depurados_final.csv a la List "Proveedores".
 *
 * - Dedupe por NIT
 * - Expande la columna "zona" con nuevas choices si aparecen en el CSV
 * - Inserta en lotes usando el endpoint $batch de Graph (20 ops por batch)
 *
 * Opciones:
 *   --dry-run   No escribe, solo reporta.
 *   --limpiar   Elimina todos los items de la List antes de migrar (destructivo).
 */

require('dotenv').config();
const fs       = require('fs');
const path     = require('path');
const g        = require('../graphStorage');
const esquemas = require('./esquemas');

const DRY_RUN  = process.argv.includes('--dry-run');
const LIMPIAR  = process.argv.includes('--limpiar');
const CSV_PATH = path.join(__dirname, '../../data/proveedores_depurados_final.csv');

// ── CSV parser mínimo con soporte de comillas ─────────────────────────────────
function parseCsv(text) {
  const rows = [];
  let   row  = [];
  let   cur  = '';
  let   inQ  = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i+1] === '"') { cur += '"'; i++; }
      else if (c === '"')                 { inQ = false; }
      else                                 { cur += c; }
    } else {
      if (c === '"')        { inQ = true; }
      else if (c === ',')   { row.push(cur); cur = ''; }
      else if (c === '\n')  { row.push(cur); rows.push(row); row = []; cur = ''; }
      else if (c === '\r')  { /* skip */ }
      else                  { cur += c; }
    }
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function limpiarZona(zona) {
  if (!zona) return null;
  const z = zona.trim();
  return z || null;
}

function limpiarTexto(s, max) {
  if (s === null || s === undefined) return '';
  const t = String(s).trim();
  return max ? t.substring(0, max) : t;
}

function limpiarNit(s) {
  return String(s || '').replace(/\D/g, '').trim();
}

// ── Transformación fila CSV → item SharePoint ─────────────────────────────────
function filaAItem(header, fila) {
  const obj = {};
  header.forEach((h, i) => { obj[h.trim()] = fila[i]; });

  const nit = limpiarNit(obj['Identificacion']);
  if (!nit) return null;

  return {
    nit,
    razonSocial:     limpiarTexto(obj['Razon social'], 255).toUpperCase(),
    nombreComercial: limpiarTexto(obj['Nombre comercial'], 255),
    regimen:         limpiarTexto(obj['Regimen tributario'], 100),
    municipio:       limpiarTexto(obj['Municipio'], 200),
    direccion:       limpiarTexto(obj['Direccion'], 255),
    telefono:        limpiarTexto(obj['Telefono principal'], 100),
    correo:          limpiarTexto(obj['Correo electronico'], 200),
    zona:            limpiarZona(obj['zona']),
    banco:           limpiarTexto(obj['Entidad bancaria'], 100),
    tipoCuenta:      limpiarTexto(obj['Tipo cuenta'], 50),
    cuentaBancaria:  limpiarTexto(obj['Cuenta bancaria'], 100),
    activo:          true,
  };
}

// ── Ajuste de choices en la columna zona ──────────────────────────────────────
async function sincronizarZonas(siteId, listId, zonasEnCsv) {
  const cols   = await g.get(`/sites/${siteId}/lists/${listId}/columns`);
  const colZona = cols.value.find(c => c.name === 'zona');
  if (!colZona) throw new Error('Columna "zona" no encontrada en Proveedores');

  const actuales = (colZona.choice && colZona.choice.choices) || [];
  const faltantes = zonasEnCsv.filter(z => z && !actuales.includes(z));

  if (!faltantes.length) {
    console.log(`  ✓ Columna "zona" ya incluye todas las opciones del CSV.`);
    return;
  }

  console.log(`  → Agregando ${faltantes.length} opción(es) nuevas a "zona": ${faltantes.join(', ')}`);
  if (DRY_RUN) return;

  const nuevas = [...actuales, ...faltantes];
  await g.patch(`/sites/${siteId}/lists/${listId}/columns/${colZona.id}`, {
    choice: { ...colZona.choice, choices: nuevas },
  });
  console.log(`  ✓ Actualizadas`);
}

// ── Batch request a Graph ─────────────────────────────────────────────────────
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

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  if (!fs.existsSync(CSV_PATH)) throw new Error(`No existe ${CSV_PATH}`);

  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  console.log(`Sitio: ${site.displayName}`);

  const list = await g.getListByName(site.id, 'Proveedores');
  if (!list) throw new Error('List "Proveedores" no encontrada. Corre crear-listas.js primero.');
  console.log(`List: Proveedores (${list.id})`);

  // Leer CSV
  const raw    = fs.readFileSync(CSV_PATH, 'utf8').replace(/^\uFEFF/, '');
  const rows   = parseCsv(raw);
  const header = rows.shift();
  console.log(`CSV: ${rows.length} filas`);

  // Transformar + dedupe por NIT
  const items = [];
  const vistos = new Set();
  const zonasEnCsv = new Set();

  for (const fila of rows) {
    const it = filaAItem(header, fila);
    if (!it) continue;
    if (vistos.has(it.nit)) continue;
    vistos.add(it.nit);
    if (it.zona) zonasEnCsv.add(it.zona);
    items.push(it);
  }
  console.log(`Únicos por NIT: ${items.length}`);

  // Sincronizar choices de zona
  await sincronizarZonas(site.id, list.id, [...zonasEnCsv]);

  // Limpieza opcional
  if (LIMPIAR) {
    console.log(`→ Eliminando items existentes (--limpiar)...`);
    const actuales = await g.getListItems(site.id, list.id);
    console.log(`  ${actuales.length} a eliminar`);
    if (!DRY_RUN) {
      for (const it of actuales) {
        await g.deleteListItem(site.id, list.id, it.id);
      }
    }
  }

  // Inserción por lotes
  const LOTE = 20;
  let   ok   = 0;
  let   fail = 0;

  for (let i = 0; i < items.length; i += LOTE) {
    const lote = items.slice(i, i + LOTE);
    if (DRY_RUN) {
      console.log(`  [dry-run] Lote ${i/LOTE + 1}: ${lote.length} items`);
      continue;
    }
    const r = await insertarLote(site.id, list.id, lote);
    ok   += r.total - r.errores.length;
    fail += r.errores.length;
    console.log(`  Lote ${i/LOTE + 1} [${ok}/${items.length}]  errores: ${r.errores.length}`);
    if (r.errores.length) {
      for (const e of r.errores.slice(0, 3)) {
        console.log(`    ✗ ${e.status}: ${JSON.stringify(e.body).substring(0, 180)}`);
      }
    }
  }

  console.log(`\n✅ Migración completada. OK: ${ok}   Errores: ${fail}`);
}

main().catch(err => {
  console.error('❌ Error fatal:', err.message);
  if (err.response) console.error(JSON.stringify(err.response, null, 2));
  process.exit(1);
});
