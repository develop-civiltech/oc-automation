'use strict';
/**
 * esquemas.js
 * Definición de las SharePoint Lists que alimentan la aplicación.
 *
 * Formato: spec compatible con Microsoft Graph API → POST /sites/{id}/lists
 * Referencia: https://learn.microsoft.com/graph/api/list-create
 *
 * Tipos de columna usados:
 *   { text: { maxLength: n } }           → texto corto
 *   { text: { allowMultipleLines: true } } → texto largo (nota)
 *   { number: { decimalPlaces: 'two' } }  → número decimal
 *   { number: { decimalPlaces: 'none' } } → entero
 *   { dateTime: {} }                     → fecha/hora
 *   { boolean: {} }                      → sí/no
 *   { choice: { choices: [...] } }       → lista de opciones
 */

const ZONAS = ['Centro', 'Caribe', 'Occidente', 'Nororiente', 'Sur', 'Llanos', 'Internacional'];

const ESTADOS_REQ = ['pendiente', 'parcial', 'gestionado', 'cerrado', 'anulado'];

const ESTADOS_OC = [
  'borrador',      // creada, sin aprobar
  'aprobada',      // aprobada por usuario → entra al contador y al registro 1.3
  'anulada',       // cancelada
  'pagada',        // check de pago (1.3)
  'entregada',     // check de logística (1.3)
  'finalizada',    // ambos checks completos
];

const ESTADOS_REM = ['activa', 'anulada', 'requiere-reemplazo'];

// ─────────────────────────────────────────────────────────────────────────────
// Requerimientos
// ─────────────────────────────────────────────────────────────────────────────
const Requerimientos = {
  displayName: 'Requerimientos',
  description: 'Solicitudes recibidas por correo o creadas manualmente',
  list: { template: 'genericList' },
  indexar: ['consecutivo', 'proyecto', 'estado'],
  columns: [
    { name: 'consecutivo',     text: { maxLength: 50 }, required: true },
    { name: 'proyecto',        text: { maxLength: 200 } },
    { name: 'fechaSolicitud',  dateTime: {} },
    { name: 'solicitante',     text: { maxLength: 200 } },
    { name: 'estado',          choice: { choices: ESTADOS_REQ, displayAs: 'dropDownMenu' } },
    { name: 'origenCorreoId',  text: { maxLength: 200 } },
    { name: 'adjuntoUrl',      text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'itemsJson',       text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'bloqueadoPor',    text: { maxLength: 200 } },
    { name: 'bloqueadoHasta',  dateTime: {} },
    { name: 'notas',           text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'ocsGeneradas',    text: { allowMultipleLines: true, textType: 'plain' } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// OrdenesCompra
// ─────────────────────────────────────────────────────────────────────────────
const OrdenesCompra = {
  displayName: 'OrdenesCompra',
  description: 'Órdenes de compra (borrador → aprobada → finalizada)',
  list: { template: 'genericList' },
  indexar: ['numeroOC', 'requerimientoId', 'proyecto', 'estado', 'tipoGasto'],
  columns: [
    { name: 'numeroOC',           text: { maxLength: 50 }, required: true },
    { name: 'requerimientoId',    text: { maxLength: 50 } },
    { name: 'cotizacionId',       text: { maxLength: 100 } },
    { name: 'proveedorNit',       text: { maxLength: 50 } },
    { name: 'proveedorNombre',    text: { maxLength: 255 } },
    { name: 'proyecto',           text: { maxLength: 255 } },
    { name: 'itemsJson',          text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'subtotal',           number: { decimalPlaces: 'two' } },
    { name: 'iva',                number: { decimalPlaces: 'two' } },
    { name: 'total',              number: { decimalPlaces: 'two' } },
    { name: 'estado',             choice: { choices: ESTADOS_OC, displayAs: 'dropDownMenu' } },
    { name: 'creadoPor',          text: { maxLength: 200 } },
    { name: 'fechaCreacion',      dateTime: {} },
    { name: 'aprobadoPor',        text: { maxLength: 200 } },
    { name: 'fechaAprobacion',    dateTime: {} },
    { name: 'anuladoPor',         text: { maxLength: 200 } },
    { name: 'fechaAnulacion',     dateTime: {} },
    { name: 'motivoAnulacion',    text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'pagado',             boolean: {} },
    { name: 'pagadoPor',          text: { maxLength: 200 } },
    { name: 'fechaPago',          dateTime: {} },
    { name: 'entregado',          boolean: {} },
    { name: 'entregadoPor',       text: { maxLength: 200 } },
    { name: 'fechaEntrega',       dateTime: {} },
    { name: 'xlsxUrl',            text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'pdfUrl',             text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'lugarEntrega',           text: { maxLength: 255 } },
    { name: 'requerimientoOrigen',    text: { maxLength: 50 } },
    { name: 'fechaEntregaPrevista',   dateTime: {} },
    { name: 'condicionesComerciales', text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'observaciones',          text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'tipoGasto',              text: { maxLength: 100 } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// Proveedores
// ─────────────────────────────────────────────────────────────────────────────
const Proveedores = {
  displayName: 'Proveedores',
  description: 'Catálogo de proveedores (migrado desde proveedores_depurados_final.csv)',
  list: { template: 'genericList' },
  indexar: ['nit', 'razonSocial', 'zona'],
  columns: [
    { name: 'nit',              text: { maxLength: 50 }, required: true },
    { name: 'razonSocial',      text: { maxLength: 255 } },
    { name: 'nombreComercial',  text: { maxLength: 255 } },
    { name: 'regimen',          text: { maxLength: 100 } },
    { name: 'municipio',        text: { maxLength: 200 } },
    { name: 'direccion',        text: { maxLength: 255 } },
    { name: 'telefono',         text: { maxLength: 100 } },
    { name: 'correo',           text: { maxLength: 200 } },
    { name: 'zona',             choice: { choices: ZONAS, displayAs: 'dropDownMenu' } },
    { name: 'banco',            text: { maxLength: 100 } },
    { name: 'tipoCuenta',       text: { maxLength: 50 } },
    { name: 'cuentaBancaria',   text: { maxLength: 100 } },
    { name: 'activo',           boolean: {} },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// Insumos (catálogo maestro con categorización)
// ─────────────────────────────────────────────────────────────────────────────
const Insumos = {
  displayName: 'Insumos',
  description: 'Catálogo maestro de insumos con categoría y subcategoría',
  list: { template: 'genericList' },
  indexar: ['nombre', 'nombreNormalizado', 'categoria', 'subcategoria'],
  columns: [
    { name: 'nombre',             text: { maxLength: 255 }, required: true },
    { name: 'nombreNormalizado',  text: { maxLength: 255 } },
    { name: 'categoria',          text: { maxLength: 100 } },
    { name: 'subcategoria',       text: { maxLength: 100 } },
    { name: 'unidadEstandar',     text: { maxLength: 50 } },
    { name: 'sinonimos',          text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'activo',             boolean: {} },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// Remisiones
// ─────────────────────────────────────────────────────────────────────────────
const Remisiones = {
  displayName: 'Remisiones',
  description: 'Remisiones de materiales asociadas a una o varias OC del mismo proyecto',
  list: { template: 'genericList' },
  indexar: ['numero', 'proyecto', 'estado'],
  columns: [
    { name: 'numero',              text: { maxLength: 50 }, required: true },
    { name: 'fecha',               dateTime: {} },
    { name: 'proyecto',            text: { maxLength: 255 } },
    { name: 'ocIds',               text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'ocsAsociadas',        text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'itemsJson',           text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'observaciones',       text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'responsableEntrega',  text: { maxLength: 200 } },
    { name: 'responsableRecepcion',text: { maxLength: 200 } },
    { name: 'lugarEntrega',        text: { maxLength: 255 } },
    { name: 'creadoPor',           text: { maxLength: 200 } },
    { name: 'fechaCreacion',       dateTime: {} },
    { name: 'estado',              choice: { choices: ESTADOS_REM, displayAs: 'dropDownMenu' } },
    { name: 'motivoAnulacion',     text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'alertas',             text: { allowMultipleLines: true, textType: 'plain' } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// Proyectos
// ─────────────────────────────────────────────────────────────────────────────
const Proyectos = {
  displayName: 'Proyectos',
  description: 'Catálogo de proyectos — solo los activos se muestran en los selectores',
  list: { template: 'genericList' },
  indexar: ['codigo', 'activo'],
  columns: [
    { name: 'codigo',       text: { maxLength: 100 }, required: true },
    { name: 'nombre',       text: { maxLength: 255 } },
    { name: 'tipo',         text: { maxLength: 100 } },
    { name: 'ciudad',       text: { maxLength: 100 } },
    { name: 'departamento', text: { maxLength: 100 } },
    { name: 'zona',         choice: { choices: ZONAS, displayAs: 'dropDownMenu' } },
    { name: 'activo',       boolean: {} },
    { name: 'notas',        text: { allowMultipleLines: true, textType: 'plain' } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// ConfiguracionApp (clave-valor)
// ─────────────────────────────────────────────────────────────────────────────
const ConfiguracionApp = {
  displayName: 'ConfiguracionApp',
  description: 'Configuración global (logo, datos de empresa, contador OC)',
  list: { template: 'genericList' },
  indexar: ['clave'],
  columns: [
    { name: 'clave',     text: { maxLength: 100 }, required: true },
    { name: 'valorJson', text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'descripcion', text: { maxLength: 255 } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// OrdenesServicio
// ─────────────────────────────────────────────────────────────────────────────
const ESTADOS_OS = ['borrador', 'aprobada', 'anulada', 'finalizada'];

const OrdenesServicio = {
  displayName: 'OrdenesServicio',
  description: 'Órdenes de servicio emitidas (borrador → aprobada → finalizada)',
  list: { template: 'genericList' },
  indexar: ['numeroOS', 'proyecto', 'estado'],
  columns: [
    { name: 'numeroOS',                   text: { maxLength: 50 }, required: true },
    { name: 'requerimientoId',            text: { maxLength: 50 } },
    { name: 'proyecto',                   text: { maxLength: 255 } },
    { name: 'proveedorNit',               text: { maxLength: 50 } },
    { name: 'proveedorNombre',            text: { maxLength: 255 } },
    { name: 'tipoServicio',               text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'clausulas',                  text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'ofertaEconomicaRef',         text: { maxLength: 200 } },
    { name: 'ofertaEconomicaCondiciones', text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'valor',                      number: { decimalPlaces: 'two' } },
    { name: 'iva',                        number: { decimalPlaces: 'two' } },
    { name: 'total',                      number: { decimalPlaces: 'two' } },
    { name: 'estado',                     choice: { choices: ESTADOS_OS, displayAs: 'dropDownMenu' } },
    { name: 'lugarPrestacion',            text: { maxLength: 255 } },
    { name: 'fechaInicio',                dateTime: {} },
    { name: 'fechaFin',                   dateTime: {} },
    { name: 'condicionesComerciales',     text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'observaciones',              text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'creadoPor',                  text: { maxLength: 200 } },
    { name: 'fechaCreacion',              dateTime: {} },
    { name: 'aprobadoPor',               text: { maxLength: 200 } },
    { name: 'fechaAprobacion',            dateTime: {} },
    { name: 'anuladoPor',                 text: { maxLength: 200 } },
    { name: 'fechaAnulacion',             dateTime: {} },
    { name: 'motivoAnulacion',            text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'tipoContrato',               text: { maxLength: 20 } },
    { name: 'itemsJson',                  text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'aiuA',                       number: { decimalPlaces: 'two' } },
    { name: 'aiuI',                       number: { decimalPlaces: 'two' } },
    { name: 'aiuU',                       number: { decimalPlaces: 'two' } },
    { name: 'tipoGasto',                  text: { maxLength: 100 } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// HistorialPrecios (reemplaza compras.csv)
// ─────────────────────────────────────────────────────────────────────────────
const HistorialPrecios = {
  displayName: 'HistorialPrecios',
  description: 'Historial transaccional de precios pagados (migrado desde compras.csv)',
  list: { template: 'genericList' },
  indexar: ['proyecto', 'insumo', 'nitProveedor', 'fecha'],
  columns: [
    { name: 'proyecto',        text: { maxLength: 100  }, required: true },
    { name: 'numeroCompra',    text: { maxLength: 80   } },
    { name: 'tipoCompra',      text: { maxLength: 50   } },
    { name: 'insumo',          text: { maxLength: 500  }, required: true },
    { name: 'cantidad',        number: { decimalPlaces: 'two' } },
    { name: 'precioUnitario',  number: { decimalPlaces: 'two' }, required: true },
    { name: 'valorTotal',      number: { decimalPlaces: 'two' } },
    { name: 'fecha',           text: { maxLength: 50   } },
    { name: 'nitProveedor',    text: { maxLength: 50   } },
    { name: 'nombreProveedor', text: { maxLength: 255  } },
    { name: 'estadoCompra',    text: { maxLength: 50   } },
    { name: 'formaPago',       text: { maxLength: 100  } },
    { name: 'anticipo',        number: { decimalPlaces: 'two' } },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// MovimientosInventario
// ─────────────────────────────────────────────────────────────────────────────
const ESTADOS_MOV = ['activo', 'anulado'];
const TIPOS_MOV   = ['entrada', 'salida'];

const MovimientosInventario = {
  displayName: 'MovimientosInventario',
  description: 'Entradas y salidas de almacén por proyecto e insumo',
  list: { template: 'genericList' },
  indexar: ['tipo', 'proyecto', 'ocId', 'insumo'],
  columns: [
    { name: 'tipo',           choice: { choices: TIPOS_MOV, displayAs: 'dropDownMenu' }, required: true },
    { name: 'fecha',          dateTime: {} },
    { name: 'proyecto',       text: { maxLength: 255 } },
    { name: 'ocId',           text: { maxLength: 50  } },
    { name: 'numeroOC',       text: { maxLength: 50  } },
    { name: 'insumo',         text: { maxLength: 500 } },
    { name: 'unidad',         text: { maxLength: 50  } },
    { name: 'cantidad',       number: { decimalPlaces: 'two' } },
    { name: 'precioUnitario', number: { decimalPlaces: 'two' } },
    { name: 'valorTotal',     number: { decimalPlaces: 'two' } },
    { name: 'responsable',    text: { maxLength: 200 } },
    { name: 'creadoPor',      text: { maxLength: 200 } },
    { name: 'fechaCreacion',  dateTime: {} },
    { name: 'notas',          text: { allowMultipleLines: true, textType: 'plain' } },
    { name: 'estado',         choice: { choices: ESTADOS_MOV, displayAs: 'dropDownMenu' } },
    { name: 'documentoRef',   text: { maxLength: 20 } },
    { name: 'estadoDoc',      choice: { choices: ['borrador', 'aprobado', 'anulado'], displayAs: 'dropDownMenu' } },
    { name: 'batchId',        text: { maxLength: 60 } },
  ],
};

module.exports = {
  todas: [Requerimientos, OrdenesCompra, Proveedores, Insumos, Remisiones, Proyectos, ConfiguracionApp, OrdenesServicio, HistorialPrecios, MovimientosInventario],
  Requerimientos, OrdenesCompra, Proveedores, Insumos, Remisiones, Proyectos, ConfiguracionApp, OrdenesServicio, HistorialPrecios, MovimientosInventario,
  ESTADOS_REQ, ESTADOS_OC, ESTADOS_REM, ESTADOS_OS, ESTADOS_MOV, TIPOS_MOV, ZONAS,
};
