'use strict';
/**
 * test.js
 * Prueba el sistema con los datos reales disponibles.
 * Ejecutar: node test.js
 */

// Apuntar las rutas a los archivos reales
process.env.PATH_COMPRAS     = '../data/compras.csv';
process.env.PATH_PROVEEDORES = '../data/proveedores_depurados_final.csv';
process.env.PATH_PROYECTOS   = '../data/tabla_proyectos.csv';

const { procesarCorreo } = require('./procesarCorreo');
const { consultarProveedor, cargarDatos } = require('./consultaProveedor');
const { parsearAsunto } = require('./parsearAsunto');

// ── TEST 1: Parseo de asunto ──────────────────────────────────────────────────
console.log('\n═══════════════════════════════════════════════════════');
console.log(' TEST 1 — Parseo de asunto del correo');
console.log('═══════════════════════════════════════════════════════');

const casosParseo = [
  'SOLICITUD REQUERIMIENTO 0001 20260410 MISTRAL',
  'SOLICITUD REQUERIMIENTO 0042 20260101 CT25-034 ANCLAJES SOLEI',
  'SOLICITUD REQUERIMIENTO 0005 20260415 CERREJON',
  'RE: Reunión del lunes',                          // debe ignorarse
  'SOLICITUD REQUERIMIENTO 99 20261399 POLANCO',    // fecha inválida
];

for (const a of casosParseo) {
  const r = parsearAsunto(a);
  console.log(`\nAsunto: "${a}"`);
  console.log(`  válido: ${r.valido}`);
  if (r.valido) console.log(`  → cons: ${r.consecutivo} | fecha: ${r.fechaTexto} | proyecto: "${r.proyecto}"`);
  else          console.log(`  → error: ${r.error}`);
}

// ── TEST 2: Correo sin adjunto ────────────────────────────────────────────────
console.log('\n═══════════════════════════════════════════════════════');
console.log(' TEST 2 — Correo sin adjunto');
console.log('═══════════════════════════════════════════════════════');

(async () => {
  const r2 = await procesarCorreo('SOLICITUD REQUERIMIENTO 0003 20260410 MISTRAL', null);
  console.log(`\nAcción: ${r2.accion}`);
  console.log('Asunto respuesta:', r2.asunto);
  console.log('Cuerpo:\n' + r2.cuerpo);
})();

// ── TEST 3: Consulta de proveedor con datos reales ────────────────────────────
console.log('\n═══════════════════════════════════════════════════════');
console.log(' TEST 3 — Consulta de proveedor por insumo');
console.log('═══════════════════════════════════════════════════════');

try {
  cargarDatos();

  const casosConsulta = [
    { insumo: 'BOTA DE SEGURIDAD EN MATERIAL PUNTA DE ACERO TALLA 40', proyecto: 'CT25-134 ANCLAJES MISTRAL' },
    { insumo: 'CASCO DIELECTRICO AZUL CON BARBUQUEJO RIGIDO DE 4 PUNTAS', proyecto: 'CT23-205 ESTABILIZACION DE TALUDES - CERREJON' },
    { insumo: 'INSUMO QUE NO EXISTE EN LA BASE', proyecto: 'CT25-075 ESTABILIZACION TALUDES - POLANCO' },
  ];

  for (const caso of casosConsulta) {
    console.log(`\nInsumo:   "${caso.insumo}"`);
    console.log(`Proyecto: "${caso.proyecto}"`);
    const r = consultarProveedor(caso.insumo, caso.proyecto);
    if (r.encontrado) {
      console.log(`  ✓ Proveedor: ${r.proveedor.nombre}`);
      console.log(`  ✓ NIT:       ${r.proveedor.nit}`);
      console.log(`  ✓ Municipio: ${r.proveedor.municipio} | Zona: ${r.proveedor.zona}`);
      console.log(`  ✓ Precio:    $${r.precio.toLocaleString('es-CO')}`);
      console.log(`  ✓ Última compra: ${r.fechaUltimaCompra} | Doc: ${r.documentoReferencia}`);
      console.log(`  ✓ Zona proyecto: ${r.zonaProyecto} | Filtró zona: ${r.aplicoFiltroZona}`);
      if (r.alertas.length) console.log('  Alertas:\n    ' + r.alertas.join('\n    '));
    } else {
      console.log(`  ✗ No encontrado: ${r.mensaje}`);
      if (r.alertas.length) console.log('  Alertas:\n    ' + r.alertas.join('\n    '));
    }
  }
} catch (e) {
  console.log('\n⚠ No se pudo cargar datos reales (ejecutar desde carpeta con /data):');
  console.log(' ', e.message);
}

console.log('\n═══════════════════════════════════════════════════════');
console.log(' Tests completados');
console.log('═══════════════════════════════════════════════════════\n');
