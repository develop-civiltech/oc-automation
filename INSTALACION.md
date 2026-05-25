# Guía de instalación — OC-Automation (Civiltech)

## ¿Qué es este software?

OC-Automation es la consola de gestión de Órdenes de Compra y Servicio de Civiltech.
Se instala localmente en cada equipo, pero toda la información se almacena en SharePoint
(nube corporativa). Esto significa que varios usuarios pueden ver y gestionar los mismos
requerimientos desde equipos distintos sin necesidad de sincronizar nada manualmente.

**Arquitectura:**
- El equipo central (Brayan) procesa automáticamente los correos de requerimiento.
- Los demás equipos solo acceden a la consola web para revisar y aprobar OCs/OS.

---

## Base de datos en la nube

Toda la información de precios, proveedores e insumos se almacena en SharePoint (listas de la nube corporativa):

| Lista SharePoint | Contenido |
|---|---|
| HistorialPrecios | Precios pagados por OC y cotización — fuente del Buscador de Precios |
| Proveedores | Catálogo activo de proveedores |
| Insumos | Catálogo maestro de insumos |
| Proyectos | Proyectos activos y su zona |
| OrdenesCompra / OrdenesServicio | Documentos emitidos |
| Requerimientos / Remisiones | Solicitudes y entregas |

Los archivos CSV en `data/` son respaldo de emergencia — el software los usa automáticamente
solo si no puede conectar con SharePoint. En condiciones normales, los CSV no son necesarios.

---

## Requisitos de instalación

Antes de comenzar, instalar en el equipo nuevo:

### 1. Node.js
- Ir a: https://nodejs.org
- Descargar el botón verde **"LTS"** (versión estable)
- Instalar con todas las opciones por defecto
- Verificar: abrir CMD y escribir `node --version` → debe mostrar un número (ej. v20.x.x)

### 2. Python 3
- Ir a: https://www.python.org/downloads/
- Descargar **"Download Python 3.13.x"**
- Durante la instalación: marcar la casilla **"Add Python to PATH"** ✅
- Después de instalar, abrir CMD y ejecutar:
  ```
  pip install openpyxl
  ```
- Verificar: `python --version` → debe mostrar Python 3.x.x

---

## Preparación de la copia maestra en OneDrive
*(Solo lo hace el administrador una vez — los demás usuarios saltan al Paso 1)*

La carpeta maestra es la fuente de instalación para todos los equipos. Se ubica en:
```
OneDrive - Civiltech IC\Software\oc-automation-master\
```

Para crearla o actualizarla:

1. Ir a `C:\Users\brayan\oc-automation\` en el equipo central
2. Copiar **la carpeta completa** `oc-automation` y pegarla en `OneDrive - Civiltech IC\Software\`
3. Renombrarla de `oc-automation` a `oc-automation-master`
4. Dentro de la carpeta ya copiada, **eliminar** los siguientes elementos:

   - [ ] Carpeta `node_modules\` — muy pesada (~50 MB), cada equipo la genera con `npm install`
   - [ ] Carpeta `logs\` — registros locales del equipo central, no son relevantes para otros
   - [ ] Carpeta `temp\` — archivos temporales del equipo central

Lo que **sí debe quedar** en la copia maestra:

   - [x] Carpeta `src\` — código del servidor
   - [x] Carpeta `ui\` — interfaz web
   - [x] Carpeta `data\` — plantillas Excel (`plantilla_oc.xlsx`, formato requerimiento) y CSVs de emergencia
     > Los precios, proveedores e insumos ya están en SharePoint — los CSV son fallback automático si SP no responde.
   - [x] `package.json` — lista de dependencias
   - [x] `index.js` — punto de entrada
   - [x] `.env` — configuración con credenciales corporativas ya completas (ver nota abajo)
   - [x] `.env.example` — plantilla de referencia
   - [x] `INSTALACION.md` — esta guía
   - [x] `iniciar-erp.bat` — acceso directo para abrir la consola (verifica Node.js e instala dependencias automáticamente si faltan)
   - [x] `actualizar.bat` — script de actualización
   - [x] `excluir-actualizar.txt` — archivo de soporte para las actualizaciones

> **Nota sobre el `.env`:** Las credenciales de Azure, SharePoint y Gemini son las mismas
> para todos los equipos. El nuevo usuario solo debe editar estos tres campos personales:
> `USUARIO_EMAIL`, `USUARIO_NOMBRE` y `OUTPUT_DIR` (ruta de salida local).
> `PYTHON_PATH` solo si la verificación del Paso 4 lo indica.

---

## Pasos de instalación en un equipo nuevo

### Paso 1 — Copiar el software

Desde el Explorador de archivos, ir a:
```
OneDrive - Civiltech IC\Software\oc-automation-master\
```
Copiar la carpeta completa `oc-automation-master` y pegarla en el equipo local.
Se recomienda ubicarla en la raíz del disco para evitar problemas con rutas largas:
```
C:\OC-Automation\
```

### Paso 2 — Instalar dependencias

Abrir CMD en la carpeta donde se copió el software:
1. Presionar `Win + R`, escribir `cmd`, Enter
2. Escribir: `cd C:\OC-Automation` y presionar Enter
3. Escribir: `npm install` y presionar Enter
4. Esperar a que finalice (puede tomar 1-2 minutos)

### Paso 3 — Personalizar el archivo .env

El archivo `.env` ya viene incluido en la copia maestra con todas las credenciales
corporativas completas. Solo es necesario editar los campos personales:

1. Abrir el archivo `.env` con el Bloc de notas
2. Buscar y actualizar únicamente estas líneas:

```
USUARIO_EMAIL=correo@civiltechic.com     ← reemplazar con el correo propio
USUARIO_NOMBRE=Nombre Apellido           ← reemplazar con el nombre propio
USUARIO_CARGO=Cargo del firmante         ← aparece al pie de OCs, OSs y remisiones
OUTPUT_DIR=C:\OC-Automation\output\      ← ajustar si la carpeta es diferente
```

El resto de valores (credenciales Azure, SharePoint, Gemini) ya están correctos
y no deben modificarse.

### Paso 4 — Verificar la instalación

En CMD, desde la carpeta del software, ejecutar:
```
node index.js --test
```

- Si aparece **"OK"** o similar → la instalación está correcta, continuar.
- Si dice **"Python no encontrado"**:
  1. Abrir CMD y escribir: `where python`
  2. Copiar la ruta que aparece (ej. `C:\Python313\python.exe`)
  3. Abrir `.env` y agregar: `PYTHON_PATH=C:\Python313\python.exe`
  4. Volver a correr `node index.js --test`

### Paso 5 — Iniciar la consola de gestión

Opción A — Doble clic en el archivo `iniciar-erp.bat`

Opción B — Desde CMD:
```
node src\servidor-cotizaciones.js
```

En ambos casos, abrir el navegador en: **http://localhost:3001**

La consola mostrará todos los requerimientos, OCs y OS que ya están en SharePoint.

---

## Uso diario

Para abrir la consola cada día:
- Doble clic en `iniciar-erp.bat`
- Abrir el navegador en http://localhost:3001

La ventana de CMD que se abre debe permanecer abierta mientras se use la consola.
Si se cierra, la página deja de responder.

---

## Recibir actualizaciones del software

Cuando el administrador avise de una nueva versión:
1. Cerrar la consola si está abierta (cerrar la ventana CMD)
2. Doble clic en `actualizar.bat`
3. Esperar a que finalice y volver a abrir la consola

El archivo `.env` y los datos locales **no se modifican** durante la actualización.

---

## Solución de problemas frecuentes

| Problema | Causa probable | Solución |
|---|---|---|
| "Puerto 3001 en uso" | Otra aplicación usa ese puerto | Cambiar `PUERTO_COTIZACIONES=3002` en `.env` y reabrir la consola |
| "Error de autenticación" | El CLIENT_SECRET expiró | Pedir al administrador el nuevo valor y actualizar `.env` |
| "Python no encontrado" | Python no está en el PATH | Agregar `PYTHON_PATH=ruta\a\python.exe` en `.env` |
| La página no carga | La consola CMD se cerró | Volver a ejecutar `iniciar-erp.bat` |
| Los datos no aparecen | Sin conexión a internet | Verificar conectividad; los datos están en SharePoint |
| El Buscador de Precios no muestra resultados | Sin conexión a SP o lista vacía | Verificar conectividad; el historial está en la lista `HistorialPrecios` de SharePoint |
| Los precios sugeridos en una OC están desactualizados | Cache de 60 seg activo | Esperar 1 minuto y recargar; o reiniciar la consola |

---

## Notas importantes

- **NO instalar la tarea automática** (`instalar-tarea.ps1`) en equipos adicionales.
  El procesamiento de correos de requerimiento corre únicamente en el equipo central.
  Los demás equipos solo acceden a la información ya procesada.

- El software **no instala nada en el sistema** más allá de su propia carpeta y Node.js/Python.
  Para desinstalarlo, simplemente eliminar la carpeta.

- **Scripts de migración** (`src\scripts\migrarCSV.js`, `src\scripts\migrarOC.js`): se ejecutaron
  una vez en el equipo central para poblar la lista `HistorialPrecios` con el historial histórico.
  **No ejecutar en equipos adicionales** — los datos ya están en SharePoint y se comparten
  automáticamente con todos los equipos conectados a la misma cuenta corporativa.

---

*Documento elaborado: Mayo 2026 — Civiltech Ingeniería y Construcción S.A.S.*
