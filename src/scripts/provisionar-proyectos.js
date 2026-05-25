'use strict';
/**
 * provisionar-proyectos.js
 * - Crea la lista SharePoint "Proyectos" si no existe (schema en esquemas.js)
 * - Siembra proyectos desde data/tabla_proyectos.csv (marca activo=true)
 * - Idempotente: si el proyecto ya existe por `codigo`, no lo duplica
 */

require('dotenv').config();
const path     = require('path');
const g        = require('../graphStorage');
const esquemas = require('./esquemas');
const XLSX     = require('xlsx');

async function asegurarLista(siteId) {
  const existing = await g.getListByName(siteId, 'Proyectos');
  if (existing) {
    console.log(`✓ Lista Proyectos ya existe (id: ${existing.id})`);
    return existing.id;
  }
  const { indexar, ...spec } = esquemas.Proyectos;
  console.log('→ Creando lista Proyectos...');
  const creada = await g.createList(siteId, spec);
  console.log(`✓ Creada id: ${creada.id}`);

  // Indexar columnas
  const cols = await g.get(`/sites/${siteId}/lists/${creada.id}/columns`);
  const mapa = new Map(cols.value.map(c => [c.name, c.id]));
  for (const n of indexar || []) {
    const id = mapa.get(n);
    if (!id) continue;
    try {
      await g.patch(`/sites/${siteId}/lists/${creada.id}/columns/${id}`, { indexed: true });
      console.log(`  ✓ indexada ${n}`);
    } catch (e) { console.log(`  ⚠ ${n}: ${e.message.split('\n')[0]}`); }
  }
  return creada.id;
}

function leerProyectosCSV() {
  const ruta = process.env.PATH_PROYECTOS || path.join(__dirname, '../../data/tabla_proyectos.csv');
  const wb   = XLSX.readFile(ruta, { raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

async function sembrar(siteId, listId) {
  const existentes = await g.getListItems(siteId, listId);
  const setCodigos = new Set(existentes.map(it => String(it.fields?.codigo || '').trim().toUpperCase()).filter(Boolean));
  const filas = leerProyectosCSV();
  let creados = 0, saltados = 0;
  for (const f of filas) {
    const codigo = String(f.codigo_proyecto || '').trim();
    if (!codigo) continue;
    if (setCodigos.has(codigo.toUpperCase())) { saltados++; continue; }
    try {
      await g.addListItem(siteId, listId, {
        codigo,
        nombre:       String(f.nombre || '').trim() || codigo,
        tipo:         String(f.tipo || '').trim(),
        ciudad:       String(f.ciudad || '').trim(),
        departamento: String(f.departamento || '').trim(),
        zona:         String(f.zona || '').trim() || 'Centro',
        activo:       true,
      });
      creados++;
    } catch (e) { console.warn(`  ✗ ${codigo}: ${e.message.split('\n')[0]}`); }
  }
  console.log(`✓ Siembra: ${creados} creados, ${saltados} ya existían.`);
}

async function main() {
  const host = process.env.SHAREPOINT_HOSTNAME;
  const sp   = process.env.SHAREPOINT_SITE_PATH;
  const site = await g.getSite(host, sp);
  const listId = await asegurarLista(site.id);
  await sembrar(site.id, listId);
  console.log('✅ Listo.');
}

main().catch(err => { console.error('❌', err.message); process.exit(1); });
