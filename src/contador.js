'use strict';
/**
 * contador.js
 * Contador de OC y OS respaldado por SharePoint.
 * El número solo se asigna cuando una orden pasa a estado 'aprobada'.
 *
 * API OC:
 *   siguienteNumeroSP(siteId, listOrdenesId) → int
 *   formato(numero)                          → string  (ej. "0042")
 *
 * API OS:
 *   siguienteNumeroOS(siteId, listOrdenesServicioId) → int
 *   formatoOS(numero)                                → string  (ej. "OS-0042")
 */

const g = require('./graphStorage');

// ── OC ────────────────────────────────────────────────────────────────────────

const ESTADOS_CONSUME_NUMERO = new Set(['aprobada', 'pagada', 'entregada', 'finalizada']);

function extraerNumero(numeroOC) {
  if (numeroOC == null) return null;
  const m = String(numeroOC).match(/(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : null;
}

function formato(numero) {
  const prefix = process.env.OC_PREFIX || '';
  const pad    = parseInt(process.env.OC_PAD || '4', 10);
  return prefix + String(numero).padStart(pad, '0');
}

async function siguienteNumeroSP(siteId, listOrdenesId) {
  const localDb = require('./db');
  // SQLite tiene todas las OCs sincronizadas; evita descargar toda la lista de SP
  const sqlite = localDb.getOrdenesCompra();
  const fuente = sqlite.length > 0
    ? sqlite
    : (await g.getListItems(siteId, listOrdenesId)).map(it => it.fields || {});
  let max = parseInt(process.env.CONTADOR_OC_INICIAL || '0', 10) - 1;
  for (const it of fuente) {
    if (!ESTADOS_CONSUME_NUMERO.has(it.estado)) continue;
    const n = extraerNumero(it.numeroOC);
    if (n != null && n > max) max = n;
  }
  return max + 1;
}

// ── OS ────────────────────────────────────────────────────────────────────────

const ESTADOS_CONSUME_NUMERO_OS = new Set(['aprobada', 'finalizada']);

function formatoOS(numero) {
  const prefix = process.env.OS_PREFIX || 'OS-';
  const pad    = parseInt(process.env.OS_PAD || '4', 10);
  return prefix + String(numero).padStart(pad, '0');
}

async function siguienteNumeroOS(siteId, listOrdenesServicioId) {
  const localDb = require('./db');
  const sqlite = localDb.getOrdenesServicio();
  const fuente = sqlite.length > 0
    ? sqlite
    : (await g.getListItems(siteId, listOrdenesServicioId)).map(it => it.fields || {});
  let max = parseInt(process.env.CONTADOR_OS_INICIAL || '0', 10) - 1;
  for (const it of fuente) {
    if (!ESTADOS_CONSUME_NUMERO_OS.has(it.estado)) continue;
    const n = extraerNumero(it.numeroOS);
    if (n != null && n > max) max = n;
  }
  return max + 1;
}

module.exports = {
  siguienteNumeroSP, extraerNumero, formato,
  siguienteNumeroOS, formatoOS,
};
