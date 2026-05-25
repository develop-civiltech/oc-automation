# instalar-tarea.ps1
# Registra la automatizacion de OC como tarea programada en Windows.
# Ejecutar como Administrador: clic derecho -> Ejecutar con PowerShell

$NombreTarea    = "Civiltech_AutomatizacionOC"
$DescripcionT   = "Revisa el buzon de abastecimiento y genera Ordenes de Compra"
$DirectorioBase = Split-Path -Parent $MyInvocation.MyCommand.Path
$RutaScript     = Join-Path $DirectorioBase "index.js"
$RutaLog        = Join-Path $DirectorioBase "logs\oc-automation.log"
$RutaLogError   = Join-Path $DirectorioBase "logs\oc-error.log"

Write-Host ""
Write-Host "=== Instalador - Civiltech OC Automation ===" -ForegroundColor Cyan
Write-Host ""

# Buscar Node.js sin operador ?. (compatible con PS 5.x)
$NodeCmd = Get-Command node -ErrorAction SilentlyContinue
if (-not $NodeCmd) {
    Write-Host "ERROR: Node.js no esta instalado." -ForegroundColor Red
    Write-Host "Descargalo en: https://nodejs.org" -ForegroundColor Yellow
    pause
    exit 1
}
$RutaNode = $NodeCmd.Source
Write-Host "Node.js encontrado: $RutaNode" -ForegroundColor Green

if (-not (Test-Path $RutaScript)) {
    Write-Host "ERROR: No se encontro index.js en $DirectorioBase" -ForegroundColor Red
    pause
    exit 1
}
Write-Host "Script encontrado: $RutaScript" -ForegroundColor Green

$RutaEnv = Join-Path $DirectorioBase ".env"
if (-not (Test-Path $RutaEnv)) {
    Write-Host "ADVERTENCIA: No se encontro el archivo .env" -ForegroundColor Yellow
}

$CarpetaLogs = Join-Path $DirectorioBase "logs"
if (-not (Test-Path $CarpetaLogs)) {
    New-Item -ItemType Directory -Path $CarpetaLogs | Out-Null
    Write-Host "Carpeta de logs creada." -ForegroundColor Green
}

# Accion: node index.js con redireccion de logs
$Argumento = """$RutaScript"" >> ""$RutaLog"" 2>> ""$RutaLogError"""
$Accion = New-ScheduledTaskAction `
    -Execute          $RutaNode `
    -Argument         $Argumento `
    -WorkingDirectory $DirectorioBase

# Disparador: L-V a las 6am, cada 5 min por 13 horas (hasta 7pm)
$DispBase = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday `
    -At "06:00AM"

$DispBase.Repetition.Interval = "PT5M"
$DispBase.Repetition.Duration = "PT13H"

$Configuracion = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit      (New-TimeSpan -Minutes 4) `
    -MultipleInstances       IgnoreNew `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -WakeToRun:$false

if (Get-ScheduledTask -TaskName $NombreTarea -ErrorAction SilentlyContinue) {
    Write-Host "Tarea existente - reemplazando..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $NombreTarea -Confirm:$false
}

Register-ScheduledTask `
    -TaskName    $NombreTarea `
    -Description $DescripcionT `
    -Action      $Accion `
    -Trigger     $DispBase `
    -Settings    $Configuracion `
    -RunLevel    Highest `
    -Force | Out-Null

Write-Host ""
Write-Host "Tarea registrada: $NombreTarea" -ForegroundColor Green
Write-Host ""
Write-Host "--- Configuracion ---" -ForegroundColor Cyan
Write-Host "  Horario:  Lunes-Viernes, 6:00 AM - 7:00 PM, cada 5 minutos"
Write-Host "  Script:   $RutaScript"
Write-Host "  Logs:     $CarpetaLogs"
Write-Host ""
Write-Host "PROXIMOS PASOS:" -ForegroundColor Yellow
Write-Host "  1. Completar .env con TENANT_ID, CLIENT_ID, CLIENT_SECRET"
Write-Host "  2. Crear buzon abastecimiento@civiltechic.com en Microsoft 365"
Write-Host "  3. Grant Admin Consent en Azure AD"
Write-Host "  4. Prueba: node index.js --test"
Write-Host "  5. Verificar en Programador de Tareas: $NombreTarea"
Write-Host ""
pause