'use strict';
/**
 * consultaProveedor.js
 * Dado un insumo y un proyecto, busca el proveedor óptimo.
 *
 * La función principal acepta datos pre-cargados (historial + proveedores desde SP)
 * o hace fallback a CSV local cuando no se pasan.
 *
 *   1. Filtrar historial de precios por insumo
 *   2. Priorizar proveedores de la misma zona del proyecto
 *   3. De los candidatos, tomar las 3 compras más recientes
 *   4. Elegir la de menor precio entre esas 3
 */

const XLSX = require('xlsx');
const path = require('path');
const fs   = require('fs');

const PATHS = {
  compras:     process.env.PATH_COMPRAS     || path.join(__dirname, '../data/compras.csv'),
  proveedores: process.env.PATH_PROVEEDORES || path.join(__dirname, '../data/proveedores_depurados_final.csv'),
  proyectos:   process.env.PATH_PROYECTOS   || path.join(__dirname, '../data/tabla_proyectos.csv'),
};

// ── Utilidades ────────────────────────────────────────────────────────────────

function parseMoney(val) {
  if (!val && val !== 0) return 0;
  return parseFloat(String(val).replace(/[^0-9.]/g, '')) || 0;
}

function parseDate(val) {
  if (!val) return null;
  const meses = {
    enero:1, febrero:2, marzo:3, abril:4, mayo:5, junio:6,
    julio:7, agosto:8, septiembre:9, octubre:10, noviembre:11, diciembre:12,
  };
  const txt = String(val).trim().toLowerCase();
  const m = txt.match(/^(\w+)\s+(\d+),?\s+(\d{4})$/);
  if (m) return new Date(+m[3], (meses[m[1]] || 1) - 1, +m[2]);
  // Formato es-CO usado al aprobar OCs: "23 de junio de 2026"
  const mEs = txt.match(/^(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})$/);
  if (mEs) return new Date(+mEs[3], (meses[mEs[2]] || 1) - 1, +mEs[1]);
  if (/^\d{4}-\d{2}-\d{2}/.test(String(val))) return new Date(String(val).slice(0, 10));
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

function normalizar(txt) {
  return String(txt || '').trim().toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// ── Fallback CSV ──────────────────────────────────────────────────────────────

function loadCSV(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    const wb = XLSX.read(content, { type: 'string' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: '' });
  } catch { return []; }
}

function cargarDatosCSV() {
  const compras     = loadCSV(PATHS.compras);
  const proveedores = loadCSV(PATHS.proveedores);
  const proyectos   = loadCSV(PATHS.proyectos);

  const provPorNit = {};
  for (const p of proveedores) {
    const nit = String(p['Identificacion'] || '').trim().replace(/\.0$/, '');
    if (nit) provPorNit[nit] = {
      nit,
      nombre:    p['Razon social'] || '',
      zona:      p['zona'] || '',
      municipio: p['Municipio'] || '',
      telefono:  p['Telefono principal'] || '',
      correo:    p['Correo electronico'] || '',
    };
  }

  const proyPorCodigo = {};
  for (const p of proyectos) {
    const cod = normalizar(p['codigo_proyecto'] || '');
    if (cod) proyPorCodigo[cod] = { zona: p['zona'] || '' };
  }

  const comprasParsed = compras.map(c => ({
    insumo:    normalizar(c['Suministro']),
    insumoRaw: String(c['Suministro'] || '').trim(),
    precio:    parseMoney(c['Valor Unitario ($)']),
    nit:       String(c['Tercero'] || '').match(/^(\d+)/)?.[1] || '',
    fecha:     parseDate(c['Fecha']),
    fechaRaw:  String(c['Fecha'] || '').trim(),
    cantidad:  parseMoney(c['Cantidad']),
    compra:    String(c['Compra'] || '').trim(),
    proyecto:  String(c['Proyecto'] || '').trim(),
    prov:      null,
  }));
  for (const c of comprasParsed) {
    if (c.nit && provPorNit[c.nit]) c.prov = provPorNit[c.nit];
  }

  return { comprasParsed, provPorNit, proyPorCodigo };
}

// ── Normalizar datos SP al mismo formato interno ──────────────────────────────

function normalizarHistorialSP(historialSP, proveedoresSP) {
  const provPorNit = {};
  for (const p of (proveedoresSP || [])) {
    if (p.nit) provPorNit[p.nit] = p;
  }

  const comprasParsed = (historialSP || []).map(h => {
    // Acepta nombres de campo de SharePoint (nitProveedor, precioUnitario, numeroCompra)
    // y nombres de campo de SQLite (nit, precio, documento)
    const nit = String(h.nitProveedor || h.nit || '').trim();
    return {
      insumo:    normalizar(h.insumo),
      insumoRaw: String(h.insumo || '').trim(),
      precio:    parseMoney(h.precioUnitario || h.precio),
      nit,
      fecha:     parseDate(h.fecha),
      fechaRaw:  String(h.fecha || '').trim(),
      cantidad:  parseMoney(h.cantidad),
      compra:    String(h.numeroCompra || h.documento || '').trim(),
      proyecto:  String(h.proyecto || '').trim(),
      prov:      provPorNit[nit] || null,
    };
  });

  return { comprasParsed, provPorNit };
}

// ── Consulta principal ────────────────────────────────────────────────────────
// opts.historialSP   → array de items de HistorialPrecios (fields de SP)
// opts.proveedoresSP → array de proveedores {nit, nombre, zona, ...}
// opts.zonaProyecto  → string de zona del proyecto
// Si no se pasan, usa fallback CSV.

function consultarProveedor(insumo, codigoProyecto, opts = {}) {
  let comprasParsed, provPorNit, zonaProyecto;

  if (opts.historialSP && opts.historialSP.length > 0) {
    const norm = normalizarHistorialSP(opts.historialSP, opts.proveedoresSP || []);
    comprasParsed = norm.comprasParsed;
    provPorNit    = norm.provPorNit;
    zonaProyecto  = opts.zonaProyecto || '';
  } else {
    const datos   = cargarDatosCSV();
    comprasParsed = datos.comprasParsed;
    provPorNit    = datos.provPorNit;
    const proyNorm = normalizar(codigoProyecto);
    zonaProyecto  = datos.proyPorCodigo[proyNorm]?.zona || '';
  }

  return ejecutarConsulta(insumo, codigoProyecto, comprasParsed, zonaProyecto);
}

function ejecutarConsulta(insumo, codigoProyecto, comprasParsed, zonaProyecto) {
  const insumoNorm = normalizar(insumo);

  let registros    = comprasParsed.filter(c => c.insumo === insumoNorm);
  let matchParcial = false;
  if (registros.length === 0) {
    registros    = comprasParsed.filter(c => c.insumo.includes(insumoNorm) || insumoNorm.includes(c.insumo));
    matchParcial = registros.length > 0;
  }

  if (registros.length === 0) {
    return {
      insumo, codigoProyecto,
      encontrado:   false,
      sinHistorial: true,
      mensaje:      `Sin historial para "${insumo}". Incluido en OC sin precio.`,
      proveedor:    null,
      precio:       null,
      historial:    [],
      alertas:      [`🔍 Ítem nuevo: "${insumo}" — complete el precio antes de aprobar la OC.`],
    };
  }

  const enriquecidos = registros
    .filter(c => c.nit && c.prov && c.precio > 0)
    .map(c => ({ ...c, zonaProveedor: normalizar(c.prov.zona || '') }));

  if (enriquecidos.length === 0) {
    return {
      insumo, codigoProyecto,
      encontrado: false, sinHistorial: false,
      mensaje:    'Proveedores históricos no están en la base activa.',
      proveedor:  null, precio: null,
      historial:  registros.slice(0, 5).map(fmtHistorial),
      alertas:    ['⚠️ Proveedores no encontrados en base depurada. Verificar manualmente.'],
    };
  }

  let candidatos       = enriquecidos;
  let aplicoFiltroZona = false;
  const zonaNorm       = normalizar(zonaProyecto);
  if (zonaNorm) {
    const enZona = enriquecidos.filter(c => c.zonaProveedor === zonaNorm);
    if (enZona.length > 0) { candidatos = enZona; aplicoFiltroZona = true; }
  }

  candidatos.sort((a, b) => (b.fecha || 0) - (a.fecha || 0));
  const top3    = candidatos.slice(0, 3);
  const elegido = [...top3].sort((a, b) => a.precio - b.precio)[0];

  const historial = enriquecidos
    .sort((a, b) => (b.fecha || 0) - (a.fecha || 0))
    .slice(0, 10)
    .map(fmtHistorial);

  const alertas = [];
  if (matchParcial)
    alertas.push(`ℹ️ Coincidencia aproximada: se usó "${elegido.insumoRaw}" para "${insumo}".`);
  if (!aplicoFiltroZona && zonaNorm)
    alertas.push(`⚠️ Sin proveedores en zona "${zonaProyecto}". Se usó historial nacional.`);
  if (top3.length === 1)
    alertas.push('ℹ️ Solo un proveedor histórico. Considere cotizar alternativas.');
  if (top3.length > 1) {
    const variacion = (Math.max(...top3.map(c => c.precio)) - Math.min(...top3.map(c => c.precio)))
                    /  Math.min(...top3.map(c => c.precio));
    if (variacion > 0.2)
      alertas.push(`📊 Variación de ${(variacion * 100).toFixed(0)}% entre proveedores recientes. Verifique antes de aprobar.`);
  }

  return {
    insumo, codigoProyecto,
    encontrado: true, sinHistorial: false,
    coincidenciaUsada: matchParcial ? elegido.insumoRaw : null,
    aplicoFiltroZona,
    zonaProyecto: zonaProyecto || 'No definida',
    proveedor: {
      nit:       elegido.nit,
      nombre:    elegido.prov?.nombre || '',
      municipio: elegido.prov?.municipio || '',
      zona:      elegido.prov?.zona || '',
      telefono:  elegido.prov?.telefono || '',
      correo:    elegido.prov?.correo || '',
    },
    precio:              elegido.precio,
    fechaUltimaCompra:   elegido.fechaRaw,
    documentoReferencia: elegido.compra,
    historial,
    alertas,
  };
}

function fmtHistorial(c) {
  return {
    nit:       c.nit || '',
    proveedor: c.prov?.nombre || c.nit || '',
    fecha:     c.fechaRaw,
    precio:    c.precio,
    cantidad:  c.cantidad,
    proyecto:  c.proyecto,
    compra:    c.compra,
  };
}

function invalidarCache() { /* no-op — cache en servidor-cotizaciones.js */ }

module.exports = { consultarProveedor, invalidarCache, cargarDatos: cargarDatosCSV };
