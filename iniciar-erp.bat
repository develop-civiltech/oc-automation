@echo off
REM iniciar-erp.bat
REM Inicia el ERP de Civiltech (OC-Automation).
REM Doble clic para abrir. Cierra la ventana para detenerla.

title Civiltech - ERP

cd /d "%~dp0"

REM Verificar Node.js
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js no está instalado.
    echo Descárgalo en https://nodejs.org
    pause
    exit /b 1
)

REM Instalar dependencias si faltan
if not exist node_modules (
    echo Instalando dependencias por primera vez...
    call npm install
)

echo.
echo ============================================
echo   Civiltech - ERP
echo ============================================
echo.
echo  La app esta iniciando...
echo  Se abrira automaticamente en tu navegador.
echo.
echo  Para cerrar la app: cierra esta ventana.
echo ============================================
echo.

REM Esperar 2 segundos y abrir el navegador
timeout /t 2 /nobreak >nul
start "" http://localhost:3001

REM Iniciar servidor
node src/servidor-cotizaciones.js

pause
