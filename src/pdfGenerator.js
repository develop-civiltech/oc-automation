'use strict';
/**
 * pdfGenerator.js
 * Renderiza HTML a PDF (Buffer) usando Chromium headless vía Puppeteer.
 */

const puppeteer = require('puppeteer');

async function htmlAPdf(html) {
  // --no-sandbox: el contenedor Docker corre como usuario no-root sin los
  // privilegios de kernel que el sandbox de Chrome necesita.
  const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    return await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '15mm', bottom: '15mm', left: '10mm', right: '10mm' },
    });
  } finally {
    await browser.close();
  }
}

module.exports = { htmlAPdf };
