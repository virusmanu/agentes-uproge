# Crea en el Escritorio el acceso directo "Agentes UPROGE" que abre la ventana de bienvenida (y quita el acceso anterior).
$raiz = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$pyw = (Get-Command pyw.exe -ErrorAction SilentlyContinue).Source
$argumentos = "-3 `"$raiz\programa\ventana.pyw`""
if (-not $pyw) { $pyw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source; $argumentos = "`"$raiz\programa\ventana.pyw`"" }
if (-not $pyw) { Write-Host "No encontré Python (pyw.exe ni pythonw.exe). Ejecute primero 1_INSTALAR.bat o instale Python."; exit 1 }
$escritorio = [Environment]::GetFolderPath('Desktop')
$ws = New-Object -ComObject WScript.Shell
$anterior = "$escritorio\Agente Ejecución MINAGRI.lnk"
if (Test-Path $anterior) { Remove-Item $anterior -ErrorAction SilentlyContinue }
$lnk = $ws.CreateShortcut("$escritorio\Agentes UPROGE.lnk")
$lnk.TargetPath = $pyw
$lnk.Arguments = $argumentos
$lnk.WorkingDirectory = "$raiz\programa"
$lnk.IconLocation = "$raiz\recursos\agente.ico,0"
$lnk.Description = "Agentes de Apoyo a la Gestión - UPROGE, Ministerio de Agricultura"
$lnk.Save()
Write-Host "Listo: se creó el acceso directo 'Agentes UPROGE' en el Escritorio ($escritorio)."
