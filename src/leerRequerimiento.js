'use strict';
/**
 * leerRequerimiento.js  v2
 * Lee el Excel adjunto CT-ADMIN-FO-002 y extrae cabecera e ítems.
 *
 * Estructura real de la hoja "Requerimientos":
 *   B8  = Proyecto   J8  = Fecha
 *   B9  = Responsable  J9 = Cargo
 *   Fila 11 = encabezados (ITEM, INSUMO, CANT, UND, NECESIDAD, POSIBLE PROVEEDOR)
 *   Filas 12-21 = ítems (B=item, C=insumo, H=cant, I=und, J=necesidad, L=posible proveedor)
 */

const XLSX = require('xlsx');

function leerRequerimiento(rutaExcel) {
  let wb;
  try {
    wb = XLSX.readFile(rutaExcel, { raw: false, cellDates: true });
  } catch (e) {
    throw new Error(`No se pudo abrir el archivo adjunto: ${e.message}`);
  }

  const ws = wb.Sheets['Requerimientos'];
  if (!ws) throw new Error('El archivo no contiene la hoja "Requerimientos". Verifique que sea el formato CT-ADMIN-FO-002.');

  const get = (celda) => {
    const c = ws[celda];
    if (!c) return '';
    const val = c.v !== undefined ? c.v : (c.w || '');
    return String(val).trim();
  };

  // Cabecera en filas 8 y 9
  let proyecto = '';
  for (const col of ['C','D','E','F','G','H','I']) {
    const v = get(`${col}8`);
    if (v) { proyecto = v; break; }
  }
  // Si no está en fila 8, intentar fila 6 (algunas versiones del formato)
  if (!proyecto) {
    for (const col of ['C','D','E','F','G','H']) {
      const v = get(`${col}6`);
      if (v && !v.includes('Versión') && !v.includes('Documento') && !v.includes('Fecha')) {
        proyecto = v; break;
      }
    }
  }

  let responsable = '';
  for (const col of ['C','D','E','F','G','H']) {
    const v = get(`${col}9`);
    if (v) { responsable = v; break; }
  }

  const cabecera = {
    proyecto,
    fecha:       get('J8') || get('K8') || '',
    responsable,
    cargo:       get('J9') || get('K9') || '',
  };

  // Buscar fila de inicio de ítems dinámicamente
  // Busca la fila donde B contiene un número (1, 2, 3...) y C tiene texto de insumo
  let filaInicio = 12; // default para V3
  for (let f = 8; f <= 20; f++) {
    const bVal = get(`B${f}`);
    const cVal = get(`C${f}`);
    // Si B tiene un número y C tiene texto, es la primera fila de ítems
    if (!isNaN(parseInt(bVal)) && cVal && !['ITEM','INSUMO','DESCRIPCION'].includes(cVal.toUpperCase())) {
      filaInicio = f;
      break;
    }
  }

  // Leer ítems: parar cuando la columna insumo (C) esté vacía
  const items = [];
  for (let fila = filaInicio; ; fila++) {
    const insumo = get(`C${fila}`);
    if (!insumo || insumo.toUpperCase() === 'INSUMO') break;

    const cantRaw  = get(`H${fila}`);
    const cantidad = isNaN(parseFloat(cantRaw)) ? 1 : parseFloat(cantRaw);
    const unidad   = get(`I${fila}`);

    if (!insumo) break;

    items.push({
      item:             get(`B${fila}`) || String(items.length + 1),
      insumo,
      cantidad,
      unidad:           unidad || 'UND',
      necesidad:        get(`J${fila}`) || '',
      posibleProveedor: get(`L${fila}`) || '',
    });
  }

  if (items.length === 0) {
    throw new Error('El formato de requerimiento no contiene ítems diligenciados. Verifique que la hoja "Requerimientos" esté completada con los insumos.');
  }

  return { cabecera, items };
}

module.exports = { leerRequerimiento };