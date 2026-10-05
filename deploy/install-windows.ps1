<#
.SYNOPSIS
    Richtet SharePoint Sync auf einem Windows-Server ein: Python-Umgebung, Web-App als NSSM-Dienst,
    Sync als geplante Aufgabe und Firewall-Regel fuer Traefik.

.DESCRIPTION
    Vorher von Hand in das Installationsverzeichnis kopieren:
      sync.py, web_app.py, requirements.txt, config.json, templates\, .env, sp_sync_cert.pfx
    (NICHT den Ordner .venv kopieren - er wird hier neu erzeugt.)
    In der .env muessen u.a. stehen (Vorlage .env.example):
      PUBLIC_URL=https://sync.contoso.com
      SP_TENANT_URL=https://contoso.sharepoint.com

    Das Skript kann mehrfach ausgefuehrt werden; vorhandene Einstellungen werden aktualisiert.
    Nur ASCII im Skript, damit Windows PowerShell 5.1 es unabhaengig von der Dateikodierung liest.

.EXAMPLE
    .\install-windows.ps1 -TraefikIp 192.0.2.10
#>
[CmdletBinding()]
param(
    # IP-Adresse des Traefik-Hosts - nur von dort wird Port 8000 freigegeben
    [Parameter(Mandatory = $true)]
    [string]$TraefikIp,

    [string]$InstallDir = "C:\SharePoint-Sync",

    # Pfad zu nssm.exe (oder nur "nssm", wenn es im PATH liegt)
    [string]$NssmPath = "nssm",

    # Python, mit dem die Umgebung angelegt wird (nur noetig, wenn .venv noch fehlt)
    [string]$PythonPath = "python",

    # Abstand zwischen zwei Sync-Laeufen
    [int]$SyncIntervalMinutes = 60,

    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"
$ServiceName = "SharePointSyncWeb"
$TaskName = "SharePoint Sync"
$VenvPython = Join-Path $InstallDir ".venv\Scripts\python.exe"

# --- Vorpruefungen ---
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw "Bitte in einer PowerShell 'Als Administrator' ausfuehren." }

foreach ($file in @("sync.py", "web_app.py", "requirements.txt", "config.json", ".env", "templates\base.html")) {
    if (-not (Test-Path (Join-Path $InstallDir $file))) { throw "Datei fehlt: $(Join-Path $InstallDir $file)" }
}
if (-not (Get-ChildItem $InstallDir -Filter *.pfx)) { throw "Zertifikat (*.pfx) fehlt in $InstallDir" }
if (-not (Get-Command $NssmPath -ErrorAction SilentlyContinue)) { throw "nssm nicht gefunden: $NssmPath (Parameter -NssmPath angeben)" }
if (-not (Select-String -Path (Join-Path $InstallDir ".env") -Pattern '^\s*PUBLIC_URL\s*=\s*https://' -Quiet)) {
    Write-Warning "In der .env fehlt PUBLIC_URL=https://... - ohne sie scheitert der Login hinter Traefik."
}

# --- 1. Python-Umgebung ---
if (-not (Test-Path $VenvPython)) {
    Write-Host "Lege Python-Umgebung an ..."
    & $PythonPath -m venv (Join-Path $InstallDir ".venv")
    if ($LASTEXITCODE -ne 0) { throw "Python-Umgebung konnte nicht angelegt werden (Python installiert? Parameter -PythonPath)." }
}
Write-Host "Installiere Python-Pakete ..."
$requirements = Join-Path $InstallDir "requirements.txt"
# Mit uv angelegte Umgebungen enthalten kein pip - dann uv verwenden, sonst pip nachinstallieren
& $VenvPython -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('pip') else 1)"
$hasPip = ($LASTEXITCODE -eq 0)
if (-not $hasPip -and (Get-Command uv -ErrorAction SilentlyContinue)) {
    & uv pip install --python $VenvPython -r $requirements
    if ($LASTEXITCODE -ne 0) { throw "uv pip install ist fehlgeschlagen." }
} else {
    if (-not $hasPip) {
        & $VenvPython -m ensurepip --upgrade
        if ($LASTEXITCODE -ne 0) { throw "pip konnte nicht in die Python-Umgebung installiert werden." }
    }
    & $VenvPython -m pip install --disable-pip-version-check -r $requirements
    if ($LASTEXITCODE -ne 0) { throw "pip install ist fehlgeschlagen." }
}

$logDir = Join-Path $InstallDir "logs"
New-Item -ItemType Directory -Force $logDir | Out-Null

# --- 2. Web-App als Windows-Dienst (NSSM) ---
Write-Host "Richte Dienst $ServiceName ein ..."
if (-not (Get-Service $ServiceName -ErrorAction SilentlyContinue)) {
    & $NssmPath install $ServiceName $VenvPython
    if ($LASTEXITCODE -ne 0) { throw "nssm install ist fehlgeschlagen." }
}
$webLog = Join-Path $logDir "web_app.log"
& $NssmPath set $ServiceName Application $VenvPython | Out-Null
# -X utf8: Umlaute und Sonderzeichen in Ausgaben unabhaengig von der Server-Sprache
& $NssmPath set $ServiceName AppParameters "-X utf8 web_app.py" | Out-Null
# Die App liest .env, config.json, Zertifikat und templates relativ zum Arbeitsverzeichnis
& $NssmPath set $ServiceName AppDirectory $InstallDir | Out-Null
& $NssmPath set $ServiceName DisplayName "SharePoint Sync Web-App" | Out-Null
& $NssmPath set $ServiceName Description "Web-Oberflaeche fuer den SharePoint-Sync (hinter Traefik, Port $Port)" | Out-Null
& $NssmPath set $ServiceName Start SERVICE_AUTO_START | Out-Null
& $NssmPath set $ServiceName AppStdout $webLog | Out-Null
& $NssmPath set $ServiceName AppStderr $webLog | Out-Null
& $NssmPath set $ServiceName AppRotateFiles 1 | Out-Null
& $NssmPath set $ServiceName AppRotateOnline 1 | Out-Null
& $NssmPath set $ServiceName AppRotateBytes 5242880 | Out-Null
& $NssmPath set $ServiceName AppEnvironmentExtra "PYTHONUNBUFFERED=1" | Out-Null
if ((Get-Service $ServiceName).Status -ne "Stopped") { & $NssmPath stop $ServiceName | Out-Null }
& $NssmPath start $ServiceName | Out-Null

# --- 3. Sync als geplante Aufgabe ---
Write-Host "Richte geplante Aufgabe '$TaskName' ein (alle $SyncIntervalMinutes Minuten) ..."
$action = New-ScheduledTaskAction -Execute $VenvPython -Argument "-X utf8 sync.py" -WorkingDirectory $InstallDir
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Minutes $SyncIntervalMinutes)
# IgnoreNew: laeuft ein Sync noch (der erste dauert lange), wird kein zweiter parallel gestartet
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 12)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description "SharePoint Mirror Sync (sync.py)" -Force | Out-Null

# --- 4. Firewall: Port nur fuer Traefik oeffnen ---
Write-Host "Richte Firewall-Regel ein (Port $Port nur von $TraefikIp) ..."
$ruleName = "SharePoint Sync Web-App (nur Traefik)"
Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule
New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port -RemoteAddress $TraefikIp | Out-Null

# --- Ergebnis ---
Start-Sleep -Seconds 5
$service = Get-Service $ServiceName
Write-Host ""
Write-Host "Dienst $ServiceName : $($service.Status)"
try {
    # /login muss zum Microsoft-Login umleiten; das prueft zugleich den Internetzugang des Servers
    $request = [System.Net.WebRequest]::Create("http://localhost:$Port/login")
    $request.AllowAutoRedirect = $false
    $request.Timeout = 20000
    $response = $request.GetResponse()
    $status = [int]$response.StatusCode
    $target = $response.Headers["Location"]
    $response.Close()
    if ($status -in 302, 307 -and $target -like "https://login.microsoftonline.com/*") {
        Write-Host "Web-App antwortet lokal und leitet zum Microsoft-Login um - in Ordnung."
    } else {
        Write-Warning "Web-App antwortet lokal unerwartet (HTTP $status) -> siehe $webLog"
    }
} catch {
    Write-Warning "Web-App antwortet lokal nicht: $($_.Exception.Message)  -> siehe $webLog"
}
Write-Host "Erster Sync-Lauf startet in ca. 5 Minuten, danach alle $SyncIntervalMinutes Minuten."
