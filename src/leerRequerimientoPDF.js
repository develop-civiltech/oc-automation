'use strict';
/**
 * leerRequerimientoPDF.js
 * Extrae {cabecera, items} desde un PDF del formato CT-ADMIN-FO-002
 * usando Gemini (vision). Devuelve la misma estructura que leerRequerimiento.js
 * para que el resto del pipeline (procesarCorreo) no cambie.
 */

const fs    = require('fs');
const https = require('https');

const GEMINI_KEY = process.env.GEMINI_API_KEY || '';
const MODELO     = 'gemini-2.5-flash';

const PROMPT = `Este PDF es el formato oficial de SOLICITUD DE REQUERIMIENTO (CT-ADMIN-FO-002) de Civiltech.
Extrae la información y devuelve SOLO un JSON (sin markdown, sin comentarios) con esta estructura EXACTA:

{
  "cabecera": {
    "proyecto":    "código o nombre exacto del proyecto",
    "fecha":       "DD/MM/YYYY (fecha de la solicitud)",
    "responsable": "nombre completo de quien solicita",
    "cargo":       "cargo del responsable"
  },
  "items": [
    {
      "item":             "1",
      "insumo":           "descripción exacta del insumo tal como aparece",
      "cantidad":         1,
      "unidad":           "UND | PAR | KG | M | GL | ...",
      "necesidad":        "fecha de necesidad o texto descriptivo si existe",
      "posibleProveedor": "proveedor sugerido si aparece, si no deja vacío"
    }
  ]
}

Reglas:
- Respeta mayúsculas/minúsculas y tildes del insumo original.
- Si la cantidad viene como texto (p.ej. "cinco"), conviértela a número.
- Si no hay unidad explícita usa "UND".
- Incluye SOLO filas con insumo y cantidad válidos — omite filas vacías o de encabezado.
- Si un campo no aparece en el PDF, déjalo como cadena vacía "".
- Devuelve al menos 1 ítem si el formato está diligenciado. Si no hay ítems, devuelve items: [].`;

async function leerRequerimientoPDF(rutaPDF) {
  if (!GEMINI_KEY) throw new Error('GEMINI_API_KEY no configurada en .env');
  if (!fs.existsSync(rutaPDF)) throw new Error(`No existe el archivo: ${rutaPDF}`);

  const pdfBase64 = fs.readFileSync(rutaPDF).toString('base64');

  const body = JSON.stringify({
    contents: [{
      parts: [
        { inline_data: { mime_type: 'application/pdf', data: pdfBase64 } },
        { text: PROMPT },
      ],
    }],
    generationConfig: { temperature: 0, maxOutputTokens: 4096 },
  });

  const resp = await new Promise((resolve, reject) => {
    const req = https.request(
      `https://generativelanguage.googleapis.com/v1beta/models/${MODELO}:generateContent?key=${GEMINI_KEY}`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' } },
      (res) => {
        const chunks = [];
        res.on('data', c => chunks.push(c));
        res.on('end', () => {
          try { resolve(JSON.parse(Buffer.concat(chunks).toString())); }
          catch (e) { reject(e); }
        });
      });
    req.on('error', reject);
    req.write(body);
    req.end();
  });

  if (resp.error) throw new Error(`Gemini: ${resp.error.message}`);

  const texto = resp.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
  const limpio = texto.replace(/```json|```/g, '').trim();

  let data;
  try { data = JSON.parse(limpio); }
  catch (e) { throw new Error(`Respuesta de Gemini no es JSON válido: ${limpio.slice(0, 200)}`); }

  const cabecera = {
    proyecto:    String(data.cabecera?.proyecto    || '').trim(),
    fecha:       String(data.cabecera?.fecha       || '').trim(),
    responsable: String(data.cabecera?.responsable || '').trim(),
    cargo:       String(data.cabecera?.cargo       || '').trim(),
  };

  const items = (Array.isArray(data.items) ? data.items : [])
    .map((it, i) => ({
      item:             String(it.item || (i + 1)),
      insumo:           String(it.insumo || '').trim(),
      cantidad:         Number(it.cantidad) || 0,
      unidad:           String(it.unidad || 'UND').trim().toUpperCase(),
      necesidad:        String(it.necesidad || '').trim(),
      posibleProveedor: String(it.posibleProveedor || '').trim(),
    }))
    .filter(it => it.insumo && it.cantidad > 0);

  if (items.length === 0) {
    throw new Error('El PDF no contiene ítems diligenciados. Verifique que el formato esté completo.');
  }

  return { cabecera, items };
}

module.exports = { leerRequerimientoPDF };
