'use strict';
/**
 * crear-control-costos.js
 * Genera el libro Excel "Control Costos.xlsx" con estructura de:
 *   - Gastos (tabla principal con todas las OC aprobadas)
 *   - Por Proyecto, Por Proveedor, Por Tipo Gasto (resúmenes con fórmulas)
 *   - Resumen (totales globales)
 *
 * Luego lo sube al sitio SharePoint Gestion-de-Proyectos en:
 *   /Documentos/Control Costos/Control Costos.xlsx
 *
 * Diseño: la hoja "Gastos" contiene una Tabla llamada "tblGastos" que
 * permite al runtime hacer POST .../tables/tblGastos/rows/add por cada
 * OC aprobada (ver graphStorage.appendRowToTable).
 */

require('dotenv').config();
const path    = require('path');
const fs      = require('fs');
const ExcelJS = require('exceljs');
const g       = require('../graphStorage');

const DRY_RUN = process.argv.includes('--dry-run');
const RECREAR = process.argv.includes('--recrear');

const RUTA_LOCAL  = path.join(__dirname, '../../temp/Control_Costos.xlsx');
const NOMBRE_REMOTO = 'Control Costos.xlsx';
const CARPETA_REMOTA = 'Control Costos';

// ── Construcción del libro ────────────────────────────────────────────────────

const HEADERS_GASTOS = [
  'Fecha OC', 'Número OC', 'Proyecto', 'Proveedor NIT', 'Proveedor Nombre',
  'Tipo Gasto', 'Subtotal', 'IVA', 'Total', 'Estado',
  'Fecha Aprobación', 'Fecha Pago', 'Fecha Entrega', 'Creado Por',
];

function estilarEncabezado(row) {
  row.height = 22;
  row.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  row.fill = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FF1E293B' },
  };
  row.alignment = { vertical: 'middle', horizontal: 'left' };
  row.border = {
    bottom: { style: 'medium', color: { argb: 'FF334155' } },
  };
}

function crearHojaGastos(wb) {
  const hoja = wb.addWorksheet('Gastos', { views: [{ state: 'frozen', ySplit: 1 }] });

  hoja.columns = [
    { header: 'Fecha OC',         key: 'fechaOC',         width: 14 },
    { header: 'Número OC',        key: 'numeroOC',        width: 12 },
    { header: 'Proyecto',         key: 'proyecto',        width: 36 },
    { header: 'Proveedor NIT',    key: 'proveedorNit',    width: 16 },
    { header: 'Proveedor Nombre', key: 'proveedorNombre', width: 40 },
    { header: 'Tipo Gasto',       key: 'tipoGasto',       width: 18 },
    { header: 'Subtotal',         key: 'subtotal',        width: 14, style: { numFmt: '$#,##0.00' } },
    { header: 'IVA',              key: 'iva',             width: 12, style: { numFmt: '$#,##0.00' } },
    { header: 'Total',            key: 'total',           width: 14, style: { numFmt: '$#,##0.00' } },
    { header: 'Estado',           key: 'estado',          width: 14 },
    { header: 'Fecha Aprobación', key: 'fechaAprobacion', width: 16 },
    { header: 'Fecha Pago',       key: 'fechaPago',       width: 14 },
    { header: 'Fecha Entrega',    key: 'fechaEntrega',    width: 14 },
    { header: 'Creado Por',       key: 'creadoPor',       width: 24 },
  ];

  estilarEncabezado(hoja.getRow(1));

  // Fila inicial vacía (ExcelJS requiere al menos 1 fila de datos para crear la tabla)
  hoja.addRow({
    fechaOC: null, numeroOC: '', proyecto: '', proveedorNit: '',
    proveedorNombre: '', tipoGasto: '', subtotal: 0, iva: 0, total: 0,
    estado: '', fechaAprobacion: null, fechaPago: null, fechaEntrega: null,
    creadoPor: '',
  });

  hoja.addTable({
    name: 'tblGastos',
    ref:  'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
    columns: HEADERS_GASTOS.map(h => ({ name: h, filterButton: true })),
    rows: [[
      null, '', '', '', '', '', 0, 0, 0, '', null, null, null, '',
    ]],
  });

  return hoja;
}

function crearResumenPorCampo(wb, hojaNombre, campo, campoHeader) {
  const hoja = wb.addWorksheet(hojaNombre, { views: [{ state: 'frozen', ySplit: 1 }] });

  hoja.columns = [
    { header: campoHeader,    key: 'clave',      width: 40 },
    { header: '# OCs',        key: 'cantidad',   width: 12 },
    { header: 'Subtotal',     key: 'subtotal',   width: 16, style: { numFmt: '$#,##0.00' } },
    { header: 'IVA',          key: 'iva',        width: 14, style: { numFmt: '$#,##0.00' } },
    { header: 'Total',        key: 'total',      width: 16, style: { numFmt: '$#,##0.00' } },
    { header: '% del Total',  key: 'porcentaje', width: 12, style: { numFmt: '0.00%' } },
  ];

  estilarEncabezado(hoja.getRow(1));

  // Fórmulas dinámicas basadas en tblGastos
  // Fila 2 en adelante: UNIQUE(tblGastos[campo]) + SUMIFS
  // Como ExcelJS no ejecuta fórmulas, las escribimos como texto y Excel las calcula al abrir.
  for (let i = 0; i < 30; i++) {
    const fila = 2 + i;
    const claveExpr = i === 0
      ? `=IFERROR(INDEX(UNIQUE(tblGastos[${campo}]),1),"")`
      : `=IFERROR(INDEX(UNIQUE(tblGastos[${campo}]),${i + 1}),"")`;
    hoja.getCell(`A${fila}`).value = { formula: claveExpr };
    hoja.getCell(`B${fila}`).value = { formula: `=IF(A${fila}="","",COUNTIF(tblGastos[${campo}],A${fila}))` };
    hoja.getCell(`C${fila}`).value = { formula: `=IF(A${fila}="","",SUMIF(tblGastos[${campo}],A${fila},tblGastos[Subtotal]))` };
    hoja.getCell(`D${fila}`).value = { formula: `=IF(A${fila}="","",SUMIF(tblGastos[${campo}],A${fila},tblGastos[IVA]))` };
    hoja.getCell(`E${fila}`).value = { formula: `=IF(A${fila}="","",SUMIF(tblGastos[${campo}],A${fila},tblGastos[Total]))` };
    hoja.getCell(`F${fila}`).value = { formula: `=IF(OR(A${fila}="",SUM(tblGastos[Total])=0),"",E${fila}/SUM(tblGastos[Total]))` };
  }

  return hoja;
}

function crearResumen(wb) {
  const hoja = wb.addWorksheet('Resumen', { views: [{ state: 'frozen', ySplit: 1 }] });

  hoja.columns = [
    { header: 'Métrica', key: 'metrica', width: 30 },
    { header: 'Valor',   key: 'valor',   width: 20 },
  ];
  estilarEncabezado(hoja.getRow(1));

  const filas = [
    { metrica: 'Total OCs registradas',  formula: '=COUNTA(tblGastos[Número OC])', fmt: '#,##0' },
    { metrica: 'OCs aprobadas',          formula: '=COUNTIF(tblGastos[Estado],"aprobada")', fmt: '#,##0' },
    { metrica: 'OCs pagadas',            formula: '=COUNTIF(tblGastos[Estado],"pagada")', fmt: '#,##0' },
    { metrica: 'OCs entregadas',         formula: '=COUNTIF(tblGastos[Estado],"entregada")', fmt: '#,##0' },
    { metrica: 'OCs finalizadas',        formula: '=COUNTIF(tblGastos[Estado],"finalizada")', fmt: '#,##0' },
    { metrica: 'OCs anuladas',           formula: '=COUNTIF(tblGastos[Estado],"anulada")', fmt: '#,##0' },
    { metrica: '',                       formula: null },
    { metrica: 'Subtotal acumulado',     formula: '=SUM(tblGastos[Subtotal])', fmt: '$#,##0.00' },
    { metrica: 'IVA acumulado',          formula: '=SUM(tblGastos[IVA])',      fmt: '$#,##0.00' },
    { metrica: 'TOTAL acumulado',        formula: '=SUM(tblGastos[Total])',    fmt: '$#,##0.00', destacado: true },
    { metrica: '',                       formula: null },
    { metrica: 'Proyectos con gasto',    formula: '=IFERROR(ROWS(UNIQUE(tblGastos[Proyecto])),0)',  fmt: '#,##0' },
    { metrica: 'Proveedores contratados',formula: '=IFERROR(ROWS(UNIQUE(tblGastos[Proveedor NIT])),0)', fmt: '#,##0' },
  ];

  for (const f of filas) {
    const row = hoja.addRow({ metrica: f.metrica, valor: f.formula ? { formula: f.formula } : '' });
    if (f.fmt) row.getCell('valor').numFmt = f.fmt;
    if (f.destacado) {
      row.font = { bold: true, size: 12 };
      row.getCell('valor').font = { bold: true, size: 14, color: { argb: 'FF059669' } };
    }
  }

  return hoja;
}

async function generarLibro() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'OC Automation';
  wb.created = new Date();

  crearHojaGastos(wb);
  crearResumenPorCampo(wb, 'Por Proyecto',   'Proyecto',         'Proyecto');
  crearResumenPorCampo(wb, 'Por Proveedor',  'Proveedor Nombre', 'Proveedor');
  crearResumenPorCampo(wb, 'Por Tipo Gasto', 'Tipo Gasto',       'Tipo de Gasto');
  crearResumen(wb);

  const dir = path.dirname(RUTA_LOCAL);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  await wb.xlsx.writeFile(RUTA_LOCAL);
  console.log(`✓ Libro generado: ${RUTA_LOCAL}`);
  return RUTA_LOCAL;
}

// ── Subida a SharePoint ───────────────────────────────────────────────────────
async function subirASharePoint(rutaLocal) {
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  console.log(`Sitio: ${site.displayName}`);

  const rutaRemota = `${CARPETA_REMOTA}/${NOMBRE_REMOTO}`;

  // Verificar si ya existe
  let existe = null;
  try {
    existe = await g.getDriveItemByPath(site.id, rutaRemota);
  } catch (err) {
    if (err.status !== 404) throw err;
  }

  if (existe && !RECREAR) {
    console.log(`⚠  Ya existe ${rutaRemota} en SharePoint (id: ${existe.id}).`);
    console.log(`   Usa --recrear para sobrescribir. Omitiendo subida.`);
    return existe;
  }

  const buffer = fs.readFileSync(rutaLocal);
  console.log(`→ Subiendo ${(buffer.length / 1024).toFixed(1)} KB a ${rutaRemota}...`);

  const resultado = await g.uploadFileToSite(
    site.id, rutaRemota, buffer,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  console.log(`✓ Subido. ID: ${resultado.id}`);
  console.log(`  URL: ${resultado.webUrl}`);
  console.log(`\n   Agrega al .env:`);
  console.log(`   CONTROL_COSTOS_ITEM_ID=${resultado.id}`);

  return resultado;
}

async function main() {
  const rutaLocal = await generarLibro();

  if (DRY_RUN) {
    console.log(`[dry-run] No se sube a SharePoint. Archivo local: ${rutaLocal}`);
    return;
  }

  await subirASharePoint(rutaLocal);
}

main().catch(err => {
  console.error('❌ Error fatal:', err.message);
  if (err.response) console.error(JSON.stringify(err.response, null, 2));
  process.exit(1);
});
