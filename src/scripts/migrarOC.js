'use strict';
/**
 * migrarOC.js — Retroalimentación de OC aprobadas → lista SharePoint "HistorialPrecios"
 *
 * Lee todas las OC aprobadas/pagadas/entregadas/finalizadas de la lista "OrdenesCompra"
 * en SharePoint y vuelca sus ítems a "HistorialPrecios", omitiendo las OC ya presentes.
 *
 * Ejecución:
 *   node src/scripts/migrarOC.js
 */

require('dotenv').config({ path: require('path').join(__dirname, '../../.env') });

const g = require('../graphStorage');

const ESTADOS_INCLUIR = new Set(['aprobada', 'pagada', 'entregada', 'finalizada']);

async function ctxSP() {
  const host     = process.env.SHAREPOINT_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !sitePath) throw new Error('SHAREPOINT_HOSTNAME / SHAREPOINT_SITE_PATH no configurados en .env');
  const site = await g.getSite(host, sitePath);

  const [lstOC, lstHP] = await Promise.all([
    g.getListByName(site.id, 'OrdenesCompra'),
    g.getListByName(site.id, 'HistorialPrecios'),
  ]);
  if (!lstOC) throw new Error('Lista OrdenesCompra no encontrada en SharePoint');
  if (!lstHP) throw new Error('Lista HistorialPrecios no encontrada en SharePoint');

  return { siteId: site.id, ocId: lstOC.id, hpId: lstHP.id };
}

async function migrar() {
  console.log('Conectando con SharePoint...');
  const { siteId, ocId, hpId } = await ctxSP();

  console.log('Cargando HistorialPrecios existente...');
  const hpItems = await g.getListItems(siteId, hpId);
  const yaRegistrados = new Set(
    hpItems
      .map(i => String(i.fields?.numeroCompra || '').trim())
      .filter(Boolean)
  );
  console.log(`  → ${yaRegistrados.size} números de compra ya en HistorialPrecios.`);

  console.log('Cargando OrdenesCompra...');
  const ocItems = await g.getListItems(siteId, ocId);
  const ocAprobadas = ocItems.filter(i => ESTADOS_INCLUIR.has(String(i.fields?.estado || '').toLowerCase()));
  console.log(`  → ${ocAprobadas.length} OC en estados aprobada/pagada/entregada/finalizada.`);

  const nuevas = ocAprobadas.filter(i => !yaRegistrados.has(String(i.fields?.numeroOC || '').trim()));
  console.log(`  → ${nuevas.length} OC sin registrar en HistorialPrecios.\n`);

  if (!nuevas.length) {
    console.log('Nada que migrar. Todas las OC ya están en HistorialPrecios.');
    return;
  }

  let ocProcesadas = 0, itemsInsertados = 0, errores = 0;

  for (const item of nuevas) {
    const oc      = item.fields || {};
    const numOC   = String(oc.numeroOC   || '').trim();
    const proyecto = String(oc.proyecto  || '').trim();
    const nit      = String(oc.proveedorNit    || '').trim().replace(/\.0$/, '');
    const nombre   = String(oc.proveedorNombre || '').trim();
    const fechaRaw = oc.fechaAprobacion || oc.fechaCreacion || '';
    const fecha    = fechaRaw ? String(fechaRaw).slice(0, 10) : '';

    let items = [];
    try { items = JSON.parse(oc.itemsJson || '[]'); } catch {}

    const conPrecio = items.filter(it => Number(it.precioUnitario || it.precio || 0) > 0);
    if (!conPrecio.length) { ocProcesadas++; continue; }

    for (const it of conPrecio) {
      const base      = Number(it.precioUnitario || it.precio || 0);
      const ivaPct    = Number(it.ivaPct || 0);
      const precioFinal = base * (1 + ivaPct / 100);
      const insumo    = String(it.descripcion || it.insumo || '').trim().toUpperCase();
      const cantidad  = Number(it.cantidad || 1);

      if (!insumo || precioFinal <= 0) continue;

      try {
        await g.addListItem(siteId, hpId, {
          proyecto,
          numeroCompra:    numOC,
          tipoCompra:      'OC',
          insumo,
          cantidad,
          precioUnitario:  precioFinal,
          valorTotal:      precioFinal * cantidad,
          fecha,
          nitProveedor:    nit,
          nombreProveedor: nombre,
          estadoCompra:    'Aprobada',
          formaPago:       '',
          anticipo:        0,
        });
        itemsInsertados++;
      } catch (e) {
        errores++;
        console.warn(`  ✗ [${numOC}] ítem "${insumo}": ${e.message}`);
      }
    }

    ocProcesadas++;
    if (ocProcesadas % 20 === 0)
      console.log(`  → ${ocProcesadas}/${nuevas.length} OC procesadas, ${itemsInsertados} ítems insertados...`);
  }

  console.log(`\nRetroalimentación completada:`);
  console.log(`  OC procesadas : ${ocProcesadas}`);
  console.log(`  Ítems insertados: ${itemsInsertados}`);
  console.log(`  Errores       : ${errores}`);
}

migrar().catch(e => { console.error('Error fatal:', e.message); process.exit(1); });
