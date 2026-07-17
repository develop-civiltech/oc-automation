'use strict';
/**
 * backfill-pdf-ordenes-servicio.js
 * Genera y sube a SharePoint (carpeta OrdenesServicioPDF) el PDF de las OS que
 * ya están marcadas como pagadas pero quedaron sin pdfUrl (creadas antes de
 * que la generación automática existiera).
 *
 * Uso:
 *   node src/scripts/backfill-pdf-ordenes-servicio.js            → dry-run (solo lista)
 *   node src/scripts/backfill-pdf-ordenes-servicio.js --confirm  → ejecuta
 *   node src/scripts/backfill-pdf-ordenes-servicio.js --confirm --todos
 *       → además de las que no tienen PDF, regenera las que ya tienen uno
 */

require('dotenv').config();
const g          = require('../graphStorage');
const osTemplate = require('../osTemplate');
const configApp  = require('../configApp');
const localDb    = require('../db');
const { htmlAPdf } = require('../pdfGenerator');

const CONFIRM = process.argv.includes('--confirm');
const TODOS   = process.argv.includes('--todos');

function osDesdeFields(item) {
  const f = item.fields || item;
  let items = [];
  try { items = JSON.parse(f.itemsJson || '[]'); } catch {}
  return {
    id:                     item.id || f.id,
    numeroOS:               f.numeroOS || '',
    fecha:                  f.fechaCreacion ? new Date(f.fechaCreacion).toLocaleDateString('es-CO') : '',
    proyecto:               f.proyecto || '',
    proveedorNit:           f.proveedorNit || '',
    proveedorNombre:        f.proveedorNombre || '',
    tipoServicio:           f.tipoServicio || '',
    clausulas:              f.clausulas || '',
    ofertaEconomicaRef:     f.ofertaEconomicaRef || '',
    ofertaEconomicaCondiciones: f.ofertaEconomicaCondiciones || '',
    tipoContrato:           f.tipoContrato || 'IVA_PLENO',
    items,
    aiuA:                   Number(f.aiuA || 0),
    aiuI:                   Number(f.aiuI || 0),
    aiuU:                   Number(f.aiuU || 0),
    valor:                  Number(f.valor || 0),
    iva:                    Number(f.iva || 0),
    total:                  Number(f.total || 0),
    estado:                 f.estado || 'borrador',
    lugarPrestacion:        f.lugarPrestacion || '',
    fechaInicio:            f.fechaInicio ? new Date(f.fechaInicio).toLocaleDateString('es-CO') : '',
    fechaFin:               f.fechaFin ? new Date(f.fechaFin).toLocaleDateString('es-CO') : '',
    condicionesComerciales: f.condicionesComerciales || '',
    observaciones:          f.observaciones || '',
    tipoGasto:              f.tipoGasto || '',
    pagado:                 !!f.pagado,
    pagadoPor:              f.pagadoPor || '',
    fechaPago:              f.fechaPago ? new Date(f.fechaPago).toLocaleDateString('es-CO') : '',
    cumplido:               !!f.cumplido,
    cumplidoPor:            f.cumplidoPor || '',
    fechaCumplido:          f.fechaCumplido ? new Date(f.fechaCumplido).toLocaleDateString('es-CO') : '',
  };
}

function cfgConFirmantePorEmail(cfg, email) {
  if (!email) return cfg;
  const usuario = localDb.getUsuarioByEmail(email) || {};
  return { ...cfg, firmante: { nombre: usuario.nombre || email, cargo: usuario.cargo || '' } };
}

async function main() {
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const lst  = await g.getListByName(site.id, 'OrdenesServicio');
  if (!lst) throw new Error('Lista OrdenesServicio no existe');

  // Por si el servidor no ha corrido asegurarListaOS() todavía en este entorno
  await g.addListColumn(site.id, lst.id, { name: 'pdfUrl', text: { allowMultipleLines: true, textType: 'plain' } }).catch(() => {});

  const todas   = localDb.getOrdenesServicio();
  const pagadas = todas.filter(os => os.pagado);
  const pendientes = TODOS ? pagadas : pagadas.filter(os => !os.pdfUrl);

  console.log(`OS pagadas: ${pagadas.length}`);
  console.log(`Pendientes de PDF${TODOS ? ' (--todos: se regeneran todas)' : ''}: ${pendientes.length}`);

  if (!CONFIRM) {
    pendientes.forEach(os => console.log(`  - id ${os.id} · ${os.numeroOS || '(sin número)'} · ${os.proyecto || '(sin proyecto)'}`));
    console.log('\nDry-run. Ejecuta con --confirm para generar y subir los PDFs.');
    return;
  }

  const cfgBase = await configApp.getConfig();
  let ok = 0, fallidos = 0;
  for (const osLocal of pendientes) {
    try {
      const item = await g.getListItem(site.id, lst.id, osLocal.id);
      const os   = osDesdeFields(item);
      const cfg  = cfgConFirmantePorEmail(cfgBase, item.fields?.pagadoPor);

      const html   = osTemplate.generarHTML(os, cfg);
      const buffer = await htmlAPdf(html);
      const nombre = `${os.numeroOS || osLocal.id}_${os.proyecto || 'SIN-PROYECTO'}`.replace(/[\\/:*?"<>|]/g, '-');
      const driveItem = await g.uploadFileToSite(site.id, `/OrdenesServicioPDF/${nombre}.pdf`, buffer, 'application/pdf');
      await g.updateListItem(site.id, lst.id, osLocal.id, { pdfUrl: driveItem.webUrl });

      ok++;
      console.log(`  ✓ id ${osLocal.id} (${ok + fallidos}/${pendientes.length})`);
    } catch (e) {
      fallidos++;
      console.error(`  ✗ id ${osLocal.id}: ${e.message}`);
    }
  }
  console.log(`\nCompletado — ok: ${ok} | fallidos: ${fallidos}`);
}

main().catch(e => { console.error('Fallo crítico:', e); process.exit(1); });
