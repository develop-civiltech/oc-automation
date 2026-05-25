'use strict';
/**
 * configApp.js
 * Lectura/escritura de configuración persistente en la List "ConfiguracionApp".
 * Claves soportadas:
 *   - logo        : data-URL (image/png;base64,...) del logo de la OC
 *   - emisor      : JSON con { razonSocial, nit, direccion, ciudad, telefono, correo, web }
 *   - firmante    : JSON con { nombre, cargo }
 *   - observacionesDefault : texto con observaciones por defecto (antes condicionesDefault)
 *
 * Esquema del item: fields = { clave, valorJson, descripcion }
 *   - valorJson guarda string plano (ej. data-url) o JSON.stringify(obj)
 */

const g = require('./graphStorage');

const EMISOR_DEFAULT = {
  razonSocial: 'CIVILTECH INGENIERÍA Y CONSTRUCCIÓN S.A.S.',
  nit:         '900.807.426-3',
  direccion:   'Cra 52A No 123-50',
  ciudad:      'Bogotá D.C., Colombia',
  telefono:    '',
  correo:      '',
  web:         '',
};
const FIRMANTE_DEFAULT = {
  nombre: 'ING. BRAYAN ALEXANDER OSPINA VASQUEZ',
  cargo:  'COORDINADOR DE PROYECTOS',
};
const CONDICIONES_DEFAULT = 'Documento No: CT-ADMIN-FO-006. ' +
  'Al recibir esta orden el proveedor acepta los términos comerciales aquí descritos.';
const IVA_DEFAULT = 19;

const _cache = {}; // { siteId, listId }
let _cfgCache = null;
let _cfgCacheAt = 0;
const CFG_TTL_MS = 5 * 60 * 1000;

async function ctx() {
  if (_cache.siteId && _cache.listId) return _cache;
  const site = await g.getSite(process.env.SHAREPOINT_HOSTNAME, process.env.SHAREPOINT_SITE_PATH);
  const list = await g.getListByName(site.id, 'ConfiguracionApp');
  if (!list) throw new Error('List "ConfiguracionApp" no existe. Ejecuta crear-listas.js');
  _cache.siteId = site.id;
  _cache.listId = list.id;
  return _cache;
}

async function obtenerItem(clave) {
  const { siteId, listId } = await ctx();
  const items = await g.getListItems(siteId, listId, { filter: `fields/clave eq '${clave}'` });
  return items[0] || null;
}

async function get(clave, fallback = null) {
  try {
    const item = await obtenerItem(clave);
    if (!item) return fallback;
    const raw = item.fields?.valorJson || '';
    try { return JSON.parse(raw); }
    catch { return raw; }
  } catch {
    return fallback;
  }
}

async function set(clave, valor, descripcion = '') {
  _cfgCache = null; // invalidar cache al guardar
  const { siteId, listId } = await ctx();
  const str = typeof valor === 'string' ? valor : JSON.stringify(valor);
  const existente = await obtenerItem(clave);
  const fields = { clave, valorJson: str };
  if (descripcion) fields.descripcion = descripcion;
  if (existente) {
    return g.updateListItem(siteId, listId, existente.id, fields);
  } else {
    return g.addListItem(siteId, listId, fields);
  }
}

async function getConfig() {
  if (_cfgCache && Date.now() - _cfgCacheAt < CFG_TTL_MS) return _cfgCache;
  const [logo, emisor, firmante, observaciones, observacionesLegacy, ivaDef] = await Promise.all([
    get('logo',        null),
    get('emisor',      EMISOR_DEFAULT),
    get('firmante',    FIRMANTE_DEFAULT),
    get('observacionesDefault', null),
    get('condicionesDefault',   null),   // clave legacy — fallback si aún no se guardó con la nueva
    get('ivaDefault',  IVA_DEFAULT),
  ]);
  const iva = Number(ivaDef);
  _cfgCache = {
    logo: logo || null,
    emisor:   { ...EMISOR_DEFAULT,   ...(emisor   || {}) },
    firmante: { ...FIRMANTE_DEFAULT, ...(firmante || {}) },
    observacionesDefault: observaciones || observacionesLegacy || CONDICIONES_DEFAULT,
    ivaDefault: Number.isFinite(iva) ? iva : IVA_DEFAULT,
  };
  _cfgCacheAt = Date.now();
  return _cfgCache;
}

module.exports = {
  get, set, getConfig,
  EMISOR_DEFAULT, FIRMANTE_DEFAULT, CONDICIONES_DEFAULT, IVA_DEFAULT,
};
