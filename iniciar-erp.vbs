' iniciar-erp.vbs
' Launcher del ERP de Civiltech — ventana CMD minimizada en la barra de tareas.
' Doble clic para iniciar. Para detener: clic derecho en taskbar -> Cerrar ventana.
' Compatible con cualquier equipo Windows sin instalar software adicional.

Dim fso, oShell, sDir
Set fso    = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
sDir = fso.GetParentFolderName(WScript.ScriptFullName)

' windowStyle 2 = minimizado con foco (visible en taskbar, se puede restaurar para ver logs)
oShell.Run "cmd /c """ & sDir & "\iniciar-erp.bat""", 2, False

' Abrir el navegador despues de 3 segundos (mismo tiempo que el .bat)
WScript.Sleep 3000
oShell.Run "http://localhost:3001"
