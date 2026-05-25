# desinstalar-tarea.ps1
# Elimina la tarea programada de automatizacion de OC.
# Ejecutar como Administrador.

$NombreTarea = "Civiltech_AutomatizacionOC"

Write-Host "" 
Write-Host "=== Desinstalador - Civiltech OC Automation ===" -ForegroundColor Cyan
Write-Host ""

if (Get-ScheduledTask -TaskName $NombreTarea -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $NombreTarea -Confirm:$false
    Write-Host "Tarea eliminada correctamente: $NombreTarea" -ForegroundColor Green
} else {
    Write-Host "La tarea no existe: $NombreTarea" -ForegroundColor Yellow
}