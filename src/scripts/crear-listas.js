'use strict';
/**
 * crear-listas.js
 * Crea (o actualiza) las SharePoint Lists en el sitio Gestión-de-Proyectos.
 *
 * Prerequisitos:
 *   1. Permisos Azure AD (app-only, admin consent):
 *        Sites.Manage.All  (o Sites.FullControl.All)
 *   2. Sitio SharePoint creado: /sites/Gestion-de-Proyectos
 *   3. .env con SHAREPOINT_HOSTNAME y SHAREPOINT_SITE_PATH
 *
 * Flujo:
 *   Paso 1 → crea la List con columnas básicas (sin 'indexed', Graph no lo acepta al crear)
 *   Paso 2 → para cada columna marcada como indexable, hace un PATCH
 *              { indexed: true } sobre la columna recién creada.
 *
 * Opciones:
 *   --dry-run   Muestra el payload sin escribir.
 *   --recrear   Elimina las listas existentes antes de crear (destructivo).
 *   --solo=Nombre   Procesa solo una lista por nombre (útil para debug).
 */

require('dotenv').config();
const g        = require('../graphStorage');
const esquemas = require('./esquemas');

const DRY_RUN = process.argv.includes('--dry-run');
const RECREAR = process.argv.includes('--recrear');
const SOLO    = (process.argv.find(a => a.startsWith('--solo=')) || '').split('=')[1];

function specSinMetadatos(esquema) {
  const { indexar, ...spec } = esquema;
  return spec;
}

async function indexarColumnas(siteId, listId, nombres) {
  if (!nombres || !nombres.length) return;
  // Obtener las columnas recién creadas para mapear nombre → id
  const colsResp = await g.get(`/sites/${siteId}/lists/${listId}/columns`);
  const mapa = new Map(colsResp.value.map(c => [c.name, c.id]));

  for (const nombre of nombres) {
    const colId = mapa.get(nombre);
    if (!colId) {
      console.log(`  ⚠  columna "${nombre}" no encontrada para indexar`);
      continue;
    }
    try {
      await g.patch(`/sites/${siteId}/lists/${listId}/columns/${colId}`, { indexed: true });
      console.log(`  ✓ indexada "${nombre}"`);
    } catch (err) {
      // No fatal: Graph a veces rechaza indexar ciertos tipos (choice, multiline)
      console.log(`  ⚠  no se pudo indexar "${nombre}": ${err.message.split('\n')[0]}`);
    }
  }
}

async function procesarLista(siteId, esquema) {
  const existing = await g.getListByName(siteId, esquema.displayName);

  if (existing && RECREAR) {
    console.log(`⚠  Eliminando List existente "${esquema.displayName}" (--recrear)`);
    if (!DRY_RUN) await g.deleteList(siteId, existing.id);
  } else if (existing) {
    console.log(`✓ List "${esquema.displayName}" ya existe (id: ${existing.id}), se omite.`);
    return;
  }

  const spec = specSinMetadatos(esquema);
  console.log(`→ Creando List "${esquema.displayName}" (${spec.columns.length} columnas)...`);

  if (DRY_RUN) {
    console.log(`  [dry-run] Payload:\n${JSON.stringify(spec, null, 2)}`);
    return;
  }

  let creada;
  try {
    creada = await g.createList(siteId, spec);
    console.log(`✓ Creada. ID: ${creada.id}`);
  } catch (err) {
    console.error(`✗ Error creando "${esquema.displayName}":`, err.message);
    console.error(`   Payload enviado:\n${JSON.stringify(spec, null, 2)}`);
    throw err;
  }

  if (esquema.indexar && esquema.indexar.length) {
    console.log(`  → Indexando ${esquema.indexar.length} columna(s)...`);
    await indexarColumnas(siteId, creada.id, esquema.indexar);
  }
  console.log('');
}

async function main() {
  const hostname = process.env.SHAREPOINT_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;

  if (!hostname || !sitePath) {
    console.error('❌ Faltan SHAREPOINT_HOSTNAME y/o SHAREPOINT_SITE_PATH en .env');
    process.exit(1);
  }

  console.log(`→ Resolviendo sitio ${hostname}${sitePath}...`);
  const site = await g.getSite(hostname, sitePath);
  console.log(`✓ Sitio encontrado. ID: ${site.id}`);
  console.log(`  Nombre: ${site.displayName}\n`);

  const lista = SOLO
    ? esquemas.todas.filter(e => e.displayName === SOLO)
    : esquemas.todas;

  if (SOLO && !lista.length) {
    console.error(`❌ No existe esquema con nombre "${SOLO}".`);
    console.error(`   Disponibles: ${esquemas.todas.map(e => e.displayName).join(', ')}`);
    process.exit(1);
  }

  for (const esquema of lista) {
    await procesarLista(site.id, esquema);
  }

  console.log('✅ Proceso completado.');
  if (!process.env.SHAREPOINT_SITE_ID) {
    console.log('   Agrega al .env:');
    console.log(`   SHAREPOINT_SITE_ID=${site.id}`);
  }
}

main().catch(err => {
  console.error('\n❌ Error fatal:', err.message);
  if (err.response) console.error('   Respuesta:', JSON.stringify(err.response, null, 2));
  process.exit(1);
});
