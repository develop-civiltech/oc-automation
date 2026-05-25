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
  const items = await g.getListItems(siteId, listOrdenesId);
  let max = parseInt(process.env.CONTADOR_OC_INICIAL || '0', 10) - 1;
  for (const it of items) {
    const f = it.fields || {};
    if (!ESTADOS_CONSUME_NUMERO.has(f.estado)) continue;
    const n = extraerNumero(f.numeroOC);
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
  const items = await g.getListItems(siteId, listOrdenesServicioId);
  let max = parseInt(process.env.CONTADOR_OS_INICIAL || '0', 10) - 1;
  for (const it of items) {
    const f = it.fields || {};
    if (!ESTADOS_CONSUME_NUMERO_OS.has(f.estado)) continue;
    const n = extraerNumero(f.numeroOS);
    if (n != null && n > max) max = n;
  }
  return max + 1;
}

module.exports = {
  siguienteNumeroSP, extraerNumero, formato,
  siguienteNumeroOS, formatoOS,
};
