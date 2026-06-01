' iniciar-erp.vbs
' Launcher del ERP de Civiltech — icono en bandeja del sistema, sin ventana CMD.
' Doble clic para iniciar.
' Clic derecho en el icono de la bandeja (esquina inferior derecha) para abrir o detener.
' Compatible con cualquier equipo Windows sin instalar software adicional.

Dim fso, oShell, sDir
Set fso    = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
sDir = fso.GetParentFolderName(WScript.ScriptFullName)

' windowStyle 0 = sin ventana visible
' -ExecutionPolicy Bypass permite ejecutar el PS1 local sin restricciones de politica
oShell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sDir & "\iniciar-erp-tray.ps1""", 0, False
