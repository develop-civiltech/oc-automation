'use strict';
/**
 * migrarCSV.js — Migración única de compras.csv → Lista SharePoint "HistorialPrecios"
 *
 * Ejecución:
 *   node src/scripts/migrarCSV.js
 *
 * Requisitos: .env configurado con SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH,
 *             SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET (o SHAREPOINT_TENANT_ID)
 */

require('dotenv').config({ path: require('path').join(__dirname, '../../.env') });

const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const PATH_COMPRAS = process.env.PATH_COMPRAS || path.join(__dirname, '../../data/compras.csv');

// Reusar el cliente de Graph del servidor
const g = require('../graphStorage');

async function asegurarLista(siteId) {
  const { HistorialPrecios: schema } = require('./esquemas');
  let listId;
  try { const lst = await g.getListByName(siteId, 'HistorialPrecios'); listId = lst?.id; } catch {}
  if (listId) return listId;

  console.log('Creando lista HistorialPrecios en SharePoint...');
  try {
    const created = await g.post(`/sites/${siteId}/lists`, {
      displayName: schema.displayName,
      description: schema.description || '',
      list: schema.list,
    });
    listId = created?.id;
  } catch (e) {
    if (!String(e.message).includes('409')) throw e;
    const lst = await g.getListByName(siteId, 'HistorialPrecios');
    listId = lst?.id;
  }
  if (!listId) throw new Error('No se pudo crear la lista HistorialPrecios');

  console.log('Agregando columnas a la lista...');
  for (const col of schema.columns) {
    const { name, required, ...tipo } = col;
    const body = { name, ...tipo };
    if (required) body.required = true;
    try { await g.post(`/sites/${siteId}/lists/${listId}/columns`, body); }
    catch (e) { if (!String(e.message).includes('409')) console.warn(`  Col "${name}":`, e.message); }
  }
  return listId;
}

async function ctxSP() {
  const host     = process.env.SHAREPOINT_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !sitePath) throw new Error('SHAREPOINT_HOSTNAME / SHAREPOINT_SITE_PATH no configurados en .env');
  const site = await g.getSite(host, sitePath);
  const listId = await asegurarLista(site.id);
  return { siteId: site.id, listId };
}

function parseMoney(val) {
  return parseFloat(String(val || '').replace(/[^0-9.]/g, '')) || 0;
}

async function migrar() {
  console.log('Leyendo compras.csv...');
  if (!fs.existsSync(PATH_COMPRAS)) throw new Error(`No se encontró ${PATH_COMPRAS}`);

  const content = fs.readFileSync(PATH_COMPRAS, 'utf-8');
  const wb = XLSX.read(content, { type: 'string' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(ws, { defval: '' });

  console.log(`Total filas CSV: ${filas.length}`);
  if (!filas.length) { console.log('Nada que migrar.'); return; }

  console.log('Conectando con SharePoint...');
  const { siteId, listId } = await ctxSP();
  console.log(`Lista HistorialPrecios encontrada: ${listId}`);

  let ok = 0, err = 0;
  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i];
    const tercero = String(fila['Tercero'] || '');
    const nitMatch = tercero.match(/^(\d+)/);
    const nit = nitMatch ? nitMatch[1] : '';
    const nombre = tercero.replace(/^\d+\s*-\s*/, '').trim();

    const precio = parseMoney(fila['Valor Unitario ($)']);
    const cant   = parseMoney(fila['Cantidad']) || 1;

    try {
      await g.addListItem(siteId, listId, {
        proyecto:        String(fila['Proyecto']    || '').trim(),
        numeroCompra:    String(fila['Compra']      || '').trim(),
        tipoCompra:      String(fila['Tipo Compra'] || 'Cotización').trim(),
        insumo:          String(fila['Suministro']  || '').trim().toUpperCase(),
        cantidad:        cant,
        precioUnitario:  precio,
        valorTotal:      precio * cant,
        fecha:           String(fila['Fecha']       || '').trim(),
        nitProveedor:    nit,
        nombreProveedor: nombre,
        estadoCompra:    String(fila['Estado']      || 'Aprobada').trim(),
        formaPago:       String(fila['Forma Pago']  || '').trim(),
        anticipo:        parseMoney(fila['Anticipo ($)']),
      });
      ok++;
      if (ok % 50 === 0) console.log(`  → ${ok}/${filas.length} filas subidas...`);
    } catch (e) {
      err++;
      console.warn(`  ✗ Fila ${i + 2}: ${e.message}`);
    }
  }

  console.log(`\nMigración completada: ${ok} OK, ${err} errores.`);
}

migrar().catch(e => { console.error('Error fatal:', e.message); process.exit(1); });
