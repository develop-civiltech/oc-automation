# iniciar-erp-tray.ps1
# Inicia el servidor ERP y crea un icono en la bandeja del sistema.
# Sin ventana CMD. Sin entrada en la barra de tareas.
# Requiere: Node.js instalado. Compatible con cualquier Windows moderno.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$dir = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- Verificar Node.js ---
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    [System.Windows.Forms.MessageBox]::Show(
        "Node.js no esta instalado.`nDescargalo en https://nodejs.org",
        "ERP CivilTech", "OK", "Error") | Out-Null
    exit
}

# --- Instalar dependencias si faltan (primera ejecucion) ---
if (-not (Test-Path "$dir\node_modules")) {
    $msg = [System.Windows.Forms.MessageBox]::Show(
        "Primera ejecucion: se instalaran las dependencias de Node.js.`n`nEsto puede tardar un momento. Continuar?",
        "ERP CivilTech", "YesNo", "Question")
    if ($msg -eq "Yes") {
        Start-Process "cmd" -ArgumentList "/c npm install" -WorkingDirectory $dir -Wait -WindowStyle Normal
    } else { exit }
}

# --- Iniciar servidor Node.js sin ventana ---
$server = Start-Process "node" -ArgumentList "src\servidor-cotizaciones.js" `
    -WorkingDirectory $dir -WindowStyle Hidden -PassThru

# --- Icono en la bandeja del sistema ---
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon    = [System.Drawing.SystemIcons]::Application
$notifyIcon.Text    = "ERP CivilTech"
$notifyIcon.Visible = $true

# Globo de notificacion al iniciar
$notifyIcon.BalloonTipTitle = "ERP CivilTech"
$notifyIcon.BalloonTipText  = "El servidor esta iniciando..."
$notifyIcon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Info
$notifyIcon.ShowBalloonTip(3000)

# --- Menu contextual (clic derecho sobre el icono) ---
$menu = New-Object System.Windows.Forms.ContextMenuStrip

$itemAbrir = New-Object System.Windows.Forms.ToolStripMenuItem "Abrir ERP en navegador"
$itemAbrir.Add_Click({ Start-Process "http://localhost:3001" })

$itemSep = New-Object System.Windows.Forms.ToolStripSeparator

$itemDetener = New-Object System.Windows.Forms.ToolStripMenuItem "Detener servidor"
$itemDetener.Add_Click({
    if (-not $server.HasExited) { $server.Kill() }
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    [System.Windows.Forms.Application]::Exit()
})

$menu.Items.AddRange([System.Windows.Forms.ToolStripItem[]]@($itemAbrir, $itemSep, $itemDetener))
$notifyIcon.ContextMenuStrip = $menu

# Doble clic en el icono -> abrir navegador
$notifyIcon.Add_DoubleClick({ Start-Process "http://localhost:3001" })

# Abrir navegador automaticamente despues de 3 segundos
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 3000
$timer.Add_Tick({
    $timer.Stop()
    Start-Process "http://localhost:3001"
})
$timer.Start()

# Mantener el proceso vivo (bucle de mensajes de Windows Forms)
[System.Windows.Forms.Application]::Run()
