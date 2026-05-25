# OC-Automation — ERP de Compras y Servicios · Civiltech IC

Consola web de gestión de requerimientos, órdenes de compra y órdenes de servicio para **Civiltech Ingeniería y Construcción S.A.S.** Se ejecuta localmente en cada equipo pero toda la información se almacena en SharePoint (nube corporativa), lo que permite que varios usuarios accedan a los mismos datos sin sincronización manual.

---

## Módulos

| Módulo | Función |
|--------|---------|
| **1.1 Requerimientos** | Visualiza solicitudes de compra recibidas por correo, agrupadas por proyecto. |
| **1.2 Generar OC** | Genera órdenes de compra desde un requerimiento o desde una cotización subida manualmente (PDF, Excel, imagen). Usa Gemini AI para extraer ítems automáticamente. |
| **1.3 Registro OCs** | Historial de órdenes de compra emitidas con búsqueda, filtros, y acciones de aprobación, pago y entrega. |
| **1.4 Órdenes de Servicio** | Crea nuevas órdenes de servicio con asistencia de IA para generar el clausulado. |
| **1.5 Registro OSs** | Historial de órdenes de servicio emitidas con edición de borradores y aprobación. |

---

## Arquitectura

```
Correos (Outlook / abastecimiento@civiltechic.com)
         │
         ▼
   index.js (procesamiento automático)
         │
         ├─ leerCorreos.js       ← Lectura de buzón vía Microsoft Graph API
         ├─ leerRequerimiento.js ← Extracción de requerimientos desde Excel adjunto
         ├─ leerRequerimientoPDF.js ← Extracción desde PDF/imagen con Gemini AI
         └─ requerimientos.js    ← Escritura a lista SharePoint "Requerimientos"

Consola web (http://localhost:3001)
         │
         ▼
   servidor-cotizaciones.js
         │
         ├─ graphStorage.js      ← Wrapper Microsoft Graph API (SharePoint)
         ├─ ocTemplate.js        ← Generación de documentos OC (HTML + Excel)
         ├─ osTemplate.js        ← Generación de documentos OS (HTML + Excel)
         ├─ remisionTemplate.js  ← Generación de remisiones
         ├─ contador.js          ← Numeración consecutiva OC / OS
         ├─ configApp.js         ← Configuración de la aplicación
         └─ Gemini API           ← Extracción de ítems desde cotizaciones
```

**Base de datos:** Listas de SharePoint (fuente principal). Los archivos CSV en `data/` son respaldo de emergencia y se usan automáticamente si SharePoint no está disponible.

| Lista SharePoint | Contenido |
|-----------------|-----------|
| `HistorialPrecios` | Precios pagados por OC y cotización (Buscador de Precios) |
| `Proveedores` | Catálogo activo de proveedores |
| `Insumos` | Catálogo maestro de insumos |
| `Proyectos` | Proyectos activos y su zona |
| `OrdenesCompra` | Órdenes de compra emitidas |
| `OrdenesServicio` | Órdenes de servicio emitidas |
| `Requerimientos` | Solicitudes de compra procesadas |
| `Remisiones` | Remisiones generadas |

---

## Requisitos

| Herramienta | Versión mínima | Descarga |
|-------------|---------------|----------|
| Node.js | 18 LTS o superior | https://nodejs.org |
| Python | 3.10 o superior | https://www.python.org/downloads/ |
| openpyxl (Python) | cualquiera | `pip install openpyxl` |

> Durante la instalación de Python marcar **"Add Python to PATH"**.

---

## Instalación

```bash
# 1. Clonar o copiar la carpeta del software al equipo
#    Recomendado: C:\OC-Automation\

# 2. Instalar dependencias Node.js
npm install

# 3. Configurar el archivo de entorno
#    Editar .env con los datos personales del usuario (ver sección Variables de entorno)
```

---

## Variables de entorno

Copiar `.env.example` a `.env` y completar los valores. Las credenciales corporativas
(Azure, SharePoint, Gemini) ya vienen configuradas en la copia maestra de OneDrive —
solo es necesario actualizar los campos personales:

```env
# ── Rutas a las bases de datos (fallback CSV) ─────────────────────────────────
PATH_COMPRAS=./data/compras.csv
PATH_PROVEEDORES=./data/proveedores_depurados_final.csv
PATH_PROYECTOS=./data/tabla_proyectos.csv

# ── Microsoft Graph API ───────────────────────────────────────────────────────
TENANT_ID=<azure-tenant-id>
CLIENT_ID=<azure-client-id>
CLIENT_SECRET=<azure-client-secret>

# ── SharePoint ────────────────────────────────────────────────────────────────
SHAREPOINT_HOSTNAME=civiltechic.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/NombreSitio

# ── Buzón de requerimientos ───────────────────────────────────────────────────
MAILBOX=abastecimiento@civiltechic.com

# ── Identificación del usuario (firma de documentos y auditoría) ──────────────
USUARIO_EMAIL=correo@civiltechic.com        # ← PERSONALIZAR
USUARIO_NOMBRE=Nombre Apellido              # ← PERSONALIZAR
USUARIO_CARGO=Cargo del firmante            # ← PERSONALIZAR

# ── IA (extracción de cotizaciones) ──────────────────────────────────────────
GEMINI_API_KEY=<api-key>

# ── Numeración de documentos ──────────────────────────────────────────────────
OC_PREFIX=OC-
OC_PAD=4
OS_PREFIX=OS-
OS_PAD=4

# ── Servidor ──────────────────────────────────────────────────────────────────
PUERTO_COTIZACIONES=3001

# ── Procesamiento automático de correos ───────────────────────────────────────
POLLING_INTERVAL_MIN=5

# ── Python (solo si no está en PATH) ─────────────────────────────────────────
# PYTHON_PATH=C:\Python313\python.exe
```

---

## Uso

### Consola web (uso diario)

```bash
# Opción A — doble clic en:
iniciar-erp.bat

# Opción B — desde CMD:
node src\servidor-cotizaciones.js
```

Abrir el navegador en **http://localhost:3001**

> La ventana CMD debe permanecer abierta mientras se use la consola.

### Procesamiento automático de correos

```bash
# Ejecutar una vez (para Tarea Programada de Windows)
node index.js

# Polling continuo cada N minutos
node index.js --watch

# Modo prueba (sin conexión al buzón)
node index.js --test
```

> El procesamiento automático de correos **solo debe correr en el equipo central** (Brayan).
> Los demás equipos solo abren la consola web.

### Tarea programada de Windows

```powershell
# Instalar
.\instalar-tarea.ps1

# Desinstalar
.\desinstalar-tarea.ps1
```

### Actualizaciones

```bash
# Doble clic en:
actualizar.bat
```

El archivo `.env` y los datos locales no se modifican durante la actualización.

---

## Estructura del proyecto

```
oc-automation/
├── index.js                          ← Punto de entrada (polling / test)
├── iniciar-erp.bat                   ← Abre la consola web
├── actualizar.bat                    ← Actualiza el software
├── instalar-tarea.ps1                ← Instala tarea programada de Windows
├── desinstalar-tarea.ps1             ← Desinstala tarea programada
├── .env                              ← Configuración local (no subir a git)
├── .env.example                      ← Plantilla de configuración
├── INSTALACION.md                    ← Guía de instalación detallada
│
├── src/
│   ├── servidor-cotizaciones.js      ← Servidor web principal (puerto 3001)
│   ├── leerCorreos.js                ← Lectura de correos vía Microsoft Graph
│   ├── procesarCorreo.js             ← Orquestador de procesamiento de correos
│   ├── leerRequerimiento.js          ← Extracción desde Excel de requerimiento
│   ├── leerRequerimientoPDF.js       ← Extracción desde PDF/imagen (Gemini AI)
│   ├── requerimientos.js             ← Operaciones sobre lista Requerimientos
│   ├── ocTemplate.js                 ← Plantilla de documento OC (HTML + Excel)
│   ├── osTemplate.js                 ← Plantilla de documento OS (HTML + Excel)
│   ├── remisionTemplate.js           ← Plantilla de remisión
│   ├── generarOC.py                  ← Generación de OC en Excel con openpyxl
│   ├── graphStorage.js               ← Wrapper Microsoft Graph API
│   ├── consultaProveedor.js          ← Búsqueda de proveedor óptimo
│   ├── controlCostos.js              ← Lista de control de costos
│   ├── configApp.js                  ← Configuración persistente de la app
│   ├── contador.js                   ← Numeración consecutiva OC / OS
│   ├── parsearAsunto.js              ← Parser de asunto de correo
│   ├── rotar-logs.js                 ← Rotación de archivos de log
│   └── scripts/
│       ├── crear-listas.js           ← Crea listas en SharePoint (setup inicial)
│       ├── migrarCSV.js              ← Migra historial CSV → SharePoint
│       ├── migrarOC.js               ← Migra OCs antiguas → SharePoint
│       ├── cargar-insumos.js         ← Carga catálogo de insumos a SharePoint
│       ├── migrar-proveedores.js     ← Migra proveedores a SharePoint
│       ├── provisionar-proyectos.js  ← Crea proyectos en SharePoint
│       └── wipe-datos-prueba.js      ← Limpia datos de prueba
│
├── ui/
│   └── consola.html                  ← Interfaz web (SPA)
│
└── data/
    ├── compras.csv                   ← Historial de compras (fallback CSV)
    ├── proveedores_depurados_final.csv ← Catálogo de proveedores (fallback)
    ├── tabla_proyectos.csv           ← Lista de proyectos (fallback)
    ├── plantilla_oc.xlsx             ← Plantilla Excel para OCs
    └── CT-ADMIN-FO-002_...xlsx       ← Formato de solicitud de requerimiento
```

---

## Solución de problemas

| Problema | Causa probable | Solución |
|----------|---------------|----------|
| "Puerto 3001 en uso" | Otro proceso usa el puerto | Cambiar `PUERTO_COTIZACIONES=3002` en `.env` |
| "Error de autenticación" | CLIENT_SECRET expirado | Solicitar al administrador el nuevo valor |
| "Python no encontrado" | Python no está en PATH | Agregar `PYTHON_PATH=ruta\python.exe` en `.env` |
| La página no carga | Consola CMD cerrada | Volver a ejecutar `iniciar-erp.bat` |
| Los datos no aparecen | Sin conexión a internet | Verificar conectividad — datos en SharePoint |
| Buscador de Precios vacío | Lista `HistorialPrecios` vacía | Verificar conexión SP o ejecutar migración |
| Precios sugeridos desactualizados | Cache activo (60 seg) | Esperar 1 min y recargar, o reiniciar consola |

---

## Notas importantes

- **Tarea programada**: instalar solo en el equipo central. Los demás equipos únicamente abren la consola web.
- **Scripts de migración** (`src/scripts/`): se ejecutaron una vez para poblar SharePoint. No ejecutar en equipos adicionales.
- **Desinstalación**: eliminar la carpeta del software. No instala nada en el sistema operativo más allá de su carpeta propia.

---

*Civiltech Ingeniería y Construcción S.A.S. · Mayo 2026*
