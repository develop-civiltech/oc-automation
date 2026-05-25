'use strict';
/**
 * controlCostos.js
 * Módulo para registrar gastos en el libro "Control Costos.xlsx" de SharePoint.
 * Se llama cada vez que una OC pasa a estado "aprobada" (o cuando cambia su
 * estado de pago/entrega).
 *
 * Uso:
 *   const cc = require('./controlCostos');
 *   await cc.registrarGasto({ ... });
 *   await cc.actualizarEstado(numeroOC, 'pagada', new Date());
 */

const g = require('./graphStorage');

const NOMBRE_TABLA    = 'tblGastos';
const CARPETA_REMOTA  = 'Control Costos';
const NOMBRE_ARCHIVO  = 'Control Costos.xlsx';

async function resolverArchivo() {
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const item = await g.getDriveItemByPath(site.id, `${CARPETA_REMOTA}/${NOMBRE_ARCHIVO}`);
  return { siteId: site.id, itemId: item.id };
}

/**
 * Añade una fila a tblGastos.
 * @param {object} oc - OC aprobada con campos mapeados a las columnas del libro.
 *   { fechaOC, numeroOC, proyecto, proveedorNit, proveedorNombre, tipoGasto,
 *     subtotal, iva, total, estado, fechaAprobacion, fechaPago, fechaEntrega, creadoPor }
 */
async function registrarGasto(oc) {
  const { siteId, itemId } = await resolverArchivo();

  const fila = [
    oc.fechaOC         || new Date().toISOString().slice(0, 10),
    oc.numeroOC        || '',
    oc.proyecto        || '',
    oc.proveedorNit    || '',
    oc.proveedorNombre || '',
    oc.tipoGasto       || '',
    Number(oc.subtotal || 0),
    Number(oc.iva      || 0),
    Number(oc.total    || 0),
    oc.estado          || 'aprobada',
    oc.fechaAprobacion || new Date().toISOString().slice(0, 10),
    oc.fechaPago       || null,
    oc.fechaEntrega    || null,
    oc.creadoPor       || '',
  ];

  return g.appendRowToTable(siteId, itemId, NOMBRE_TABLA, fila);
}

/**
 * Actualiza el estado (y fechas) de una OC ya registrada en tblGastos.
 * Busca por numeroOC.
 */
async function actualizarFila(numeroOC, cambios) {
  const { siteId, itemId } = await resolverArchivo();

  // Leer la tabla para encontrar la fila
  const resp = await g.get(`/sites/${siteId}/drive/items/${itemId}/workbook/tables/${NOMBRE_TABLA}/rows`);
  const filas = resp.value || [];
  const idx = filas.findIndex(r => r.values && r.values[0] && r.values[0][1] === numeroOC);
  if (idx < 0) throw new Error(`OC "${numeroOC}" no encontrada en tblGastos`);

  const fila = filas[idx].values[0];
  // Mapeo columnas (0-index): 0=FechaOC 1=NumeroOC 2=Proyecto 3=ProvNit 4=ProvNombre
  // 5=TipoGasto 6=Subtotal 7=IVA 8=Total 9=Estado 10=FechaAprobacion 11=FechaPago
  // 12=FechaEntrega 13=CreadoPor
  const MAPA = {
    fechaOC: 0, numeroOC: 1, proyecto: 2, proveedorNit: 3, proveedorNombre: 4,
    tipoGasto: 5, subtotal: 6, iva: 7, total: 8, estado: 9,
    fechaAprobacion: 10, fechaPago: 11, fechaEntrega: 12, creadoPor: 13,
  };
  for (const [k, v] of Object.entries(cambios)) {
    if (k in MAPA) fila[MAPA[k]] = v;
  }

  return g.patch(`/sites/${siteId}/drive/items/${itemId}/workbook/tables/${NOMBRE_TABLA}/rows/itemAt(index=${idx})`, {
    values: [fila],
  });
}

module.exports = { registrarGasto, actualizarFila };
