'use strict';
/**
 * rotar-logs.js
 * Mantiene los archivos de log bajo control:
 *   - Si el log supera 5 MB lo renombra con fecha y crea uno nuevo
 *   - Elimina logs con más de 30 días de antigüedad
 *
 * Se ejecuta automáticamente al inicio de cada ciclo en index.js
 */

const fs   = require('fs');
const path = require('path');

const DIR_LOGS     = path.join(__dirname, '../logs');
const MAX_BYTES    = 5 * 1024 * 1024; // 5 MB
const MAX_DIAS     = 30;
const ARCHIVOS     = ['oc-automation.log', 'oc-error.log'];

function rotar() {
  if (!fs.existsSync(DIR_LOGS)) {
    fs.mkdirSync(DIR_LOGS, { recursive: true });
    return;
  }

  const ahora    = new Date();
  const fechaStr = ahora.toISOString().slice(0, 10); // YYYY-MM-DD

  for (const nombre of ARCHIVOS) {
    const ruta = path.join(DIR_LOGS, nombre);
    if (!fs.existsSync(ruta)) continue;

    const stat = fs.statSync(ruta);

    // Rotar si supera el tamaño máximo
    if (stat.size > MAX_BYTES) {
      const base    = nombre.replace('.log', '');
      const archivo = `${base}_${fechaStr}.log`;
      fs.renameSync(ruta, path.join(DIR_LOGS, archivo));
      fs.writeFileSync(ruta, `--- Log rotado el ${ahora.toISOString()} ---\n`);
    }
  }

  // Eliminar logs viejos
  const limite = Date.now() - MAX_DIAS * 24 * 60 * 60 * 1000;
  for (const archivo of fs.readdirSync(DIR_LOGS)) {
    if (!archivo.endsWith('.log')) continue;
    const ruta = path.join(DIR_LOGS, archivo);
    const stat = fs.statSync(ruta);
    if (stat.mtimeMs < limite) {
      fs.unlinkSync(ruta);
    }
  }
}

module.exports = { rotar };
