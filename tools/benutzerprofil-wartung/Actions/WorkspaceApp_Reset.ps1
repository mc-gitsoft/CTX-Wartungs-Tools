[CmdletBinding()]
param(
    [ValidateSet('Interactive','Silent')]
    [string]$Mode = 'Interactive',
    [object]$Params = @{}
)

<#PSParamsSchema
[
    { "name": "KillAllCitrixProcesses", "type": "bool", "default": false, "label": "Alle Citrix-Prozesse beenden (nicht nur Receiver/Workspace)" }
]
PSParamsSchema#>

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot 'shared\WartungsTools.SDK.psm1') -Force
$Params = ConvertTo-Hashtable -InputObject $Params
$toolId = (Get-Content (Join-Path $toolRoot 'tool.json') -Raw | ConvertFrom-Json).toolId
$actionName = 'WorkspaceApp_Reset'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log ("Mode={0}" -f $Mode)
Write-Log "=== Citrix Workspace App Reset gestartet ==="

# ==========================
# Prozesslisten
# ==========================
# Basis: fuer Profil-/SelfService-Reset relevante Prozesse
$baseCitrixProcesses = @(
    'SelfService','SelfServicePlugin','Receiver','wfcrun32','WebHelper',
    'AuthManager','redirector','concentr','cdviewer'
)

# Optional: zusaetzliche HDX-/Sitzungsprozesse (nur mit KillAllCitrixProcesses)
$extraCitrixProcesses = @(
    'wfica32','concentr','pnamain','wfshell','ssonsvr'
)

# ==========================
# CleanUp.exe ausfuehren
# ==========================
function Invoke-CitrixCleanup {
    Write-Log "Fuehre Citrix CleanUp.exe (cleanUser) aus..."

    $cleanUpCandidates = @(
        "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\CleanUp.exe",
        "C:\Program Files (x86)\Citrix\online plugin\ica client\SelfServicePlugin\CleanUp.exe",
        "C:\Program Files (x86)\Citrix\ICA Client\CleanUp.exe",
        "C:\Program Files\Citrix\ICA Client\SelfServicePlugin\CleanUp.exe",
        "C:\Program Files\Citrix\ICA Client\CleanUp.exe"
    )

    $cleanUpPath = $null
    foreach ($candidate in $cleanUpCandidates) {
        if (Test-Path $candidate) {
            $cleanUpPath = $candidate
            break
        }
    }

    # Fallback: unter Citrix-Verzeichnissen suchen
    if (-not $cleanUpPath) {
        $roots = @("C:\Program Files (x86)\Citrix", "C:\Program Files\Citrix") | Where-Object { Test-Path $_ }
        foreach ($r in $roots) {
            try {
                $hit = Get-ChildItem -Path $r -Filter "CleanUp.exe" -Recurse -ErrorAction SilentlyContinue |
                       Select-Object -First 1 -ExpandProperty FullName
                if ($hit) { $cleanUpPath = $hit; break }
            } catch { }
        }
    }

    if ($cleanUpPath) {
        Write-Log ("Verwende CleanUp.exe: {0}" -f $cleanUpPath)
        try {
            & $cleanUpPath -cleanUser -silent
            $exitCode = $LASTEXITCODE
            if ($null -eq $exitCode) { $exitCode = 0 }
            if ($exitCode -eq 0) {
                Write-Log "CleanUp.exe erfolgreich ausgefuehrt (ExitCode 0)."
            } else {
                Write-Log ("CleanUp.exe beendet mit ExitCode {0} (versionsabhaengig)." -f $exitCode) 'WARN'
            }
            Start-Sleep -Seconds 5
        } catch {
            Write-Log ("Fehler beim Ausfuehren der CleanUp.exe: {0}" -f $_.Exception.Message) 'ERROR'
        }
    } else {
        Write-Log "Keine CleanUp.exe unter den bekannten Pfaden gefunden." 'WARN'
    }
}

# ==========================
# Hauptfunktion
# ==========================
function Invoke-WorkspaceReset {
    [CmdletBinding()]
    param(
        [switch]$KillAllCitrixProcesses
    )

    Write-Log "Starte Citrix Workspace App Reset..."

    # 1) Prozessliste bestimmen
    if ($KillAllCitrixProcesses) {
        $targetProcesses = $baseCitrixProcesses + $extraCitrixProcesses
        Write-Log "Aggressiver Prozess-Reset (KillAllCitrixProcesses aktiv)."
    } else {
        $targetProcesses = $baseCitrixProcesses
    }

    # 2) Prozesse beenden (SDK)
    Write-Log "Beende Citrix-Prozesse des aktuellen Benutzers..."
    $killed = WartungsTools.SDK\Stop-SessionProcesses -ProcessNames $targetProcesses -Retries 10 -DelayMs 400
    Write-Log ("Citrix-Prozesse beendet ({0} gestoppt)." -f $killed)

    # 3) Benutzerordner bereinigen (SDK)
    Write-Log "Loesche Citrix-Ordner im Benutzerprofil..."

    $pathsToDelete = @(
        (Join-Path $env:APPDATA      'Citrix\SelfService'),
        (Join-Path $env:APPDATA      'Citrix\Receiver'),
        (Join-Path $env:APPDATA      'Citrix\ICA Client'),
        (Join-Path $env:LOCALAPPDATA 'Citrix\SelfService'),
        (Join-Path $env:LOCALAPPDATA 'Citrix\Receiver'),
        (Join-Path $env:LOCALAPPDATA 'Citrix\Workspace')
    )

    foreach ($p in $pathsToDelete) {
        if (Test-Path $p) {
            $removed = WartungsTools.SDK\Remove-PathSafe -Path $p
            if ($removed) { Write-Log ("Geloescht: {0}" -f $p) }
            else { Write-Log ("Konnte nicht vollstaendig geloescht werden: {0}" -f $p) 'WARN' }
        } else {
            Write-Log ("Nicht vorhanden: {0}" -f $p)
        }
    }

    # 4) HKCU-Registry bereinigen (SDK)
    Write-Log "Entferne Citrix-Registry-Eintraege (HKCU)..."

    $regKeysToDelete = @(
        'HKCU:\Software\Citrix\Receiver',
        'HKCU:\Software\Citrix\SelfService',
        'HKCU:\Software\Citrix\Dazzle',
        'HKCU:\Software\Citrix\ICA Client',
        'HKCU:\Software\Citrix\AuthManager'
    )

    foreach ($key in $regKeysToDelete) {
        $removed = WartungsTools.SDK\Clear-RegistryPath -Path $key
        if ($removed) { Write-Log ("Registry geloescht: {0}" -f $key) }
        else { Write-Log ("Registry konnte nicht geloescht werden: {0}" -f $key) 'WARN' }
    }

    # 5) CleanUp.exe ausfuehren
    Invoke-CitrixCleanup

    Write-Log "Citrix Workspace App Reset abgeschlossen."
    return @{ ExitCode = 0; Errors = 0; Warnings = 0 }
}

# ==========================
# Ausfuehrung
# ==========================
if ($Mode -eq 'Silent') {
    $killAll = $false
    if ($Params.ContainsKey('KillAllCitrixProcesses')) { $killAll = [bool]$Params.KillAllCitrixProcesses }

    $result = Invoke-WorkspaceReset -KillAllCitrixProcesses:$killAll
    exit $result.ExitCode
}

# ==========================
# Interactive: Ausfuehrung mit Benutzerausgabe
# ==========================
$result = Invoke-WorkspaceReset

Write-Host ""
Write-Host "Citrix Workspace App wurde im Benutzerprofil so gruendlich wie moeglich zurueckgesetzt." -ForegroundColor Green
Write-Host "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
Write-Host ""

Write-Log "=== Citrix Workspace App Reset beendet ==="
