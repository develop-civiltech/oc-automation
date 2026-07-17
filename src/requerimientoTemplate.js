'use strict';
/**
 * requerimientoTemplate.js
 * HTML de un Requerimiento para su PDF de respaldo (backend), con el mismo
 * formato visual que "Exportar selección" en la consola (ui/consola.html).
 */

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

function fmtFecha(f) {
  if (!f) return '';
  const d = new Date(f);
  return isNaN(d) ? '' : d.toLocaleDateString('es-CO');
}

function generarHTML(fields, items) {
  const titulo = `${fields.consecutivoSistema || fields.consecutivo || 'REQ'} — ${fields.proyecto || ''}`;

  const filas = (items || []).map((it, i) => `<tr>
    <td>${i + 1}</td>
    <td>${esc(it.insumo || '—')}</td>
    <td style="text-align:right">${it.cantidad || 0}</td>
    <td>${esc(it.unidad || 'UND')}</td>
    <td>${esc(it.necesidad || '')}</td>
    <td>${esc(it.posibleProveedor || '')}</td>
  </tr>`).join('');

  return `<!DOCTYPE html><html lang="es"><head>
    <meta charset="UTF-8"><title>${titulo}</title>
    <style>
      body{font-family:Arial,sans-serif;font-size:12px;color:#1a2430;margin:24px}
      h2{color:#6b1f2a;margin-bottom:2px}
      .meta{color:#666;margin-bottom:14px;font-size:11px}
      table{width:100%;border-collapse:collapse}
      th{background:#6b1f2a;color:#fff;padding:6px 8px;text-align:left;font-size:11px}
      td{padding:5px 8px;border-bottom:1px solid #e5e7eb;vertical-align:top}
      tr:nth-child(even){background:#f9f9f9}
    </style>
  </head><body>
    <h2>${titulo}</h2>
    <p class="meta">Fecha: ${esc(fmtFecha(fields.fechaSolicitud))} · Solicitante: ${esc(fields.solicitante || '—')}</p>
    <table>
      <thead><tr><th>#</th><th>Insumo</th><th>Solicit.</th><th>Unidad</th><th>Necesidad</th><th>Posible proveedor</th></tr></thead>
      <tbody>${filas}</tbody>
    </table>
  </body></html>`;
}

module.exports = { generarHTML };
