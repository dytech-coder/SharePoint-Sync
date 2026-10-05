<#
.SYNOPSIS
    Entfernt SharePoint Sync wieder vom Windows-Server: NSSM-Dienst, geplante Aufgabe und Firewall-Regel
    (alles, was install-windows.ps1 eingerichtet hat).

.DESCRIPTION
    Ohne -RemoveFiles bleibt das Installationsverzeichnis unberuehrt. Dort liegen dann weiterhin
    das Zertifikat mit privatem Schluessel (*.pfx) und die .env.
    Mit -RemoveFiles wird das gesamte Verzeichnis geloescht - vorher sichern, was auf den neuen
    Server umziehen soll (config.json, runs\, .env, sp_sync_cert.pfx).

    -WhatIf zeigt nur an, was entfernt wuerde.
    Nur ASCII im Skript, damit Windows PowerShell 5.1 es unabhaengig von der Dateikodierung liest.

.EXAMPLE
    .\uninstall-windows.ps1
.EXAMPLE
    .\uninstall-windows.ps1 -RemoveFiles
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$InstallDir = "C:\SharePoint-Sync",

    # Pfad zu nssm.exe (oder nur "nssm", wenn es im PATH liegt)
    [string]$NssmPath = "nssm",

    # Loescht zusaetzlich das komplette Installationsverzeichnis
    [switch]$RemoveFiles
)

$ErrorActionPreference = "Stop"
# Muessen zu install-windows.ps1 passen
$ServiceName = "SharePointSyncWeb"
$TaskName = "SharePoint Sync"
$RuleName = "SharePoint Sync Web-App (nur Traefik)"

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw "Bitte in einer PowerShell 'Als Administrator' ausfuehren." }

# --- 1. Geplante Aufgabe (zuerst, damit kein neuer Sync mehr startet) ---
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    if ($PSCmdlet.ShouldProcess("Geplante Aufgabe '$TaskName'", "Stoppen und entfernen")) {
        # Bricht einen gerade laufenden Sync ab
        Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Geplante Aufgabe '$TaskName' entfernt."
    }
} else {
    Write-Host "Geplante Aufgabe '$TaskName' nicht vorhanden."
}

# --- 2. Web-App-Dienst ---
if (Get-Service $ServiceName -ErrorAction SilentlyContinue) {
    if ($PSCmdlet.ShouldProcess("Dienst $ServiceName", "Stoppen und entfernen")) {
        if ((Get-Service $ServiceName).Status -ne "Stopped") {
            Stop-Service $ServiceName -Force
            (Get-Service $ServiceName).WaitForStatus("Stopped", (New-TimeSpan -Seconds 30))
        }
        if (Get-Command $NssmPath -ErrorAction SilentlyContinue) {
            & $NssmPath remove $ServiceName confirm | Out-Null
        } else {
            # NSSM nicht (mehr) vorhanden - der Dienst laesst sich auch mit Bordmitteln entfernen
            & sc.exe delete $ServiceName | Out-Null
        }
        if ($LASTEXITCODE -ne 0) { throw "Dienst $ServiceName konnte nicht entfernt werden." }
        Write-Host "Dienst $ServiceName entfernt."
    }
} else {
    Write-Host "Dienst $ServiceName nicht vorhanden."
}

# --- 3. Firewall-Regel ---
$rules = @(Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue)
if ($rules.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("Firewall-Regel '$RuleName'", "Entfernen")) {
        $rules | Remove-NetFirewallRule
        Write-Host "Firewall-Regel '$RuleName' entfernt."
    }
} else {
    Write-Host "Firewall-Regel '$RuleName' nicht vorhanden."
}

# --- 4. Dateien ---
if ($RemoveFiles) {
    if (-not (Test-Path $InstallDir)) {
        Write-Host "Verzeichnis $InstallDir nicht vorhanden."
    } elseif (-not (Test-Path (Join-Path $InstallDir "sync.py"))) {
        # Schutz vor einem falsch angegebenen Pfad
        throw "$InstallDir sieht nicht nach SharePoint Sync aus (sync.py fehlt) - es wird nichts geloescht."
    } elseif ($PSCmdlet.ShouldProcess($InstallDir, "Verzeichnis mit allen Dateien loeschen")) {
        # Nicht aus dem Verzeichnis heraus loeschen, das gerade entfernt wird
        Set-Location $env:SystemRoot
        Remove-Item $InstallDir -Recurse -Force
        Write-Host "Verzeichnis $InstallDir geloescht."
    }
} elseif (Test-Path $InstallDir) {
    Write-Host ""
    Write-Warning "Die Dateien in $InstallDir sind noch vorhanden - darunter das Zertifikat mit privatem Schluessel (*.pfx) und die .env."
    Write-Host "Zum Loeschen des Verzeichnisses: .\uninstall-windows.ps1 -RemoveFiles"
}
