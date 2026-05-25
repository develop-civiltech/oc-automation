@echo off
setlocal

set "MASTER=C:\Users\%USERNAME%\OneDrive - Civiltech IC\Software\oc-automation-master"

if not exist "%MASTER%" (
  echo.
  echo ERROR: No se encontro la carpeta maestra en:
  echo   %MASTER%
  echo.
  echo Verifica que OneDrive este sincronizado y que la carpeta exista.
  pause
  exit /b 1
)

echo.
echo ============================================
echo    Actualizando OC-Automation...
echo ============================================
echo.

xcopy "%MASTER%\src"  "%~dp0src\"  /E /Y /Q /I
xcopy "%MASTER%\ui"   "%~dp0ui\"   /E /Y /Q /I
xcopy "%MASTER%\*.js" "%~dp0"      /Y /Q /EXCLUDE:"%~dp0excluir-actualizar.txt"
xcopy "%MASTER%\package.json" "%~dp0" /Y /Q

echo Instalando dependencias...
call npm install --silent --prefix "%~dp0"

echo.
echo ============================================
echo   Actualizacion completada correctamente.
echo   Reinicia la consola si estaba abierta.
echo ============================================
echo.
pause
