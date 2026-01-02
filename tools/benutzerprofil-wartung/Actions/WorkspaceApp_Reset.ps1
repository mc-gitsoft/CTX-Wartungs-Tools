[CmdletBinding()]
param(
    [ValidateSet('Interactive','Silent')]
    [string]$Mode = 'Interactive',
    [hashtable]$Params = @{},
    # Optional: auch Laufzeit-/Sitzungsprozesse (wfica32 etc.) beenden.
    # WARNUNG: Beendet ggf. aktive Sitzungen dieses Benutzers in dieser Session.
    [switch]$KillAllCitrixProcesses
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot 'shared\WartungsTools.SDK.psm1') -Force
$toolId = (Get-Content (Join-Path $toolRoot 'tool.json') -Raw | ConvertFrom-Json).toolId
$actionName = 'WorkspaceApp_Reset'

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

if ($Params.ContainsKey('KillAllCitrixProcesses')) {
    $KillAllCitrixProcesses = [bool]$Params.KillAllCitrixProcesses
}
# --------------------------------------------------------------------
# 1) Helper: Citrix-Prozesse nur in der aktuellen User-Session beenden
# --------------------------------------------------------------------
function Stop-CitrixProcessesUserScoped {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ProcessNames,

        [int]$Retries = 10,
        [int]$DelayMs = 400
    )

    try {
        $currentSession = (Get-Process -Id $PID).SessionId
    }
    catch {
        Write-Log "Konnte aktuelle SessionId nicht ermitteln, beende Prozesse ohne Session-Filter." 'WARN'
        $currentSession = $null
    }

    # >>> Konsolidierung: Name-Handling robuster (case-insensitive, unique)
    $processNamesNormalized = $ProcessNames | ForEach-Object { $_.ToLower() } | Select-Object -Unique

    for ($i = 1; $i -le $Retries; $i++) {
        try {
            if ($currentSession -ne $null) {
                # Nur Prozesse in der aktuellen Session
                Get-Process -ErrorAction SilentlyContinue | Where-Object {
                    ($processNamesNormalized -contains $_.Name.ToLower()) -and $_.SessionId -eq $currentSession
                } | Stop-Process -Force -ErrorAction SilentlyContinue
            }
            else {
                # Fallback: kein Session-Filter m�glich
                Get-Process -ErrorAction SilentlyContinue | Where-Object {
                    $processNamesNormalized -contains $_.Name.ToLower()
                } | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Fehler beim Beenden von Citrix-Prozessen: $($_.Exception.Message)" 'WARN'
        }

        Start-Sleep -Milliseconds $DelayMs
    }
}

# --------------------------------------------------------------------
# 2) Prozesslisten definieren
# --------------------------------------------------------------------
# Basis: f�r Profil-/SelfService-Reset relevante Prozesse
$baseCitrixProcesses = @(
    "SelfService",
    "SelfServicePlugin",
    "Receiver",
    "wfcrun32",
    "WebHelper",
    # >>> Erg�nzungen wie gew�nscht
    "AuthManager",
    "redirector",
    "concentr",
    "cdviewer"
)

# Optional: zus�tzliche HDX-/Sitzungsprozesse (nur mit -KillAllCitrixProcesses)
$extraCitrixProcesses = @(
    "wfica32",          # ICA-Client
    "concentr",         # Connection Center (bleibt, auch wenn schon in base -> wird dedupliziert)
    "pnamain",          # pnagent / �lter
    "wfshell",          # Citrix Shell
    "ssonsvr"           # Single Sign-On Service im Userkontext
)

if ($KillAllCitrixProcesses) {
    $targetProcesses = $baseCitrixProcesses + $extraCitrixProcesses
    Write-Log "Starte aggressiven Prozess-Reset (-KillAllCitrixProcesses aktiv)." 'INFO'
} else {
    $targetProcesses = $baseCitrixProcesses
}

# --------------------------------------------------------------------
# 3) Prozesse beenden (nur aktuelle Session)
# --------------------------------------------------------------------
Write-Log "[1/4] Beende Citrix-Prozesse des aktuellen Benutzers..." 'INFO'
Stop-CitrixProcessesUserScoped -ProcessNames $targetProcesses

# --------------------------------------------------------------------
# 4) Benutzerordner bereinigen
# --------------------------------------------------------------------
Write-Log "[2/4] L�sche Citrix-Ordner im Benutzerprofil..." 'INFO'

$pathsToDelete = @(
    "$env:APPDATA\Citrix\SelfService",
    "$env:APPDATA\Citrix\Receiver",
    "$env:APPDATA\Citrix\ICA Client",
    "$env:LOCALAPPDATA\Citrix\SelfService",
    "$env:LOCALAPPDATA\Citrix\Receiver",
    "$env:LOCALAPPDATA\Citrix\Workspace"
)

foreach ($path in $pathsToDelete) {
    try {
        if (Test-Path $path) {
            Write-Log "L�sche Ordner: $path" 'INFO'
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log "Fehler beim L�schen von '$path': $($_.Exception.Message)" 'WARN'
    }
}

# --------------------------------------------------------------------
# 5) HKCU-Registry-Reste entfernen
# --------------------------------------------------------------------
Write-Log "[3/4] Entferne Citrix-Registry-Eintr�ge (HKCU)..." 'INFO'

$regKeysToDelete = @(
    "HKCU:\Software\Citrix\Receiver",
    "HKCU:\Software\Citrix\SelfService",
    "HKCU:\Software\Citrix\Dazzle",
    "HKCU:\Software\Citrix\ICA Client",
    "HKCU:\Software\Citrix\AuthManager"
)

foreach ($key in $regKeysToDelete) {
    try {
        if (Test-Path $key) {
            Write-Log "L�sche Registry-Schl�ssel: $key" 'INFO'
            Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log "Fehler beim L�schen von Registry-Schl�ssel '$key': $($_.Exception.Message)" 'WARN'
    }
}

# --------------------------------------------------------------------
# 6) Zum Schluss: CleanUp.exe (cleanUser) ausf�hren
# --------------------------------------------------------------------
Write-Log "[4/4] F�hre Citrix CleanUp.exe (cleanUser) aus..." 'INFO'

# >>> Erweiterte Kandidatenliste (32/64-bit + weitere typische Pfade)
$cleanUpCandidates = @(
    # bew�hrte Pfade (deine)
    "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\CleanUp.exe",
    "C:\Program Files (x86)\Citrix\online plugin\ica client\SelfServicePlugin\CleanUp.exe",

    # weitere h�ufige Pfade
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

# >>> Optionaler Fallback: wenn nicht gefunden, einmal kurz unter Citrix suchen (sparsam)
if (-not $cleanUpPath) {
    $roots = @("C:\Program Files (x86)\Citrix", "C:\Program Files\Citrix") | Where-Object { Test-Path $_ }
    foreach ($r in $roots) {
        try {
            $hit = Get-ChildItem -Path $r -Filter "CleanUp.exe" -Recurse -ErrorAction SilentlyContinue |
                   Select-Object -First 1 -ExpandProperty FullName
            if ($hit) { $cleanUpPath = $hit; break }
        }
        catch {
            # keine harte Fehlermeldung n�tig
        }
    }
}

if ($cleanUpPath) {
    Write-Log "Verwende CleanUp.exe: $cleanUpPath" 'INFO'
    try {
        & $cleanUpPath -cleanUser -silent

        # >>> ExitCode auswerten + loggen
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq $null) { $exitCode = 0 }  # defensive
        if ($exitCode -eq 0) {
            Write-Log "CleanUp.exe erfolgreich ausgef�hrt (ExitCode 0)." 'INFO'
        } else {
            Write-Log "CleanUp.exe beendet mit ExitCode $exitCode (versionsabh�ngig). Bitte Ergebnis pr�fen." 'WARN'
        }

        Start-Sleep -Seconds 5
    }
    catch {
        Write-Log "Fehler beim Ausf�hren der CleanUp.exe: $($_.Exception.Message)" 'ERROR'
    }
} else {
    Write-Log "Keine CleanUp.exe unter den bekannten Pfaden gefunden." 'WARN'
}

# Sicherheitshalber noch einmal kurz Prozesse abr�umen
#Stop-CitrixProcessesUserScoped -ProcessNames $targetProcesses -Retries 3 -DelayMs 300

Write-Host ""
Write-Host "Citrix Workspace App wurde im Benutzerprofil so gr�ndlich wie m�glich zur�ckgesetzt." -ForegroundColor Green
Write-Host "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
Write-Host ""
Write-Log "Citrix Workspace App Reset im Benutzerprofil abgeschlossen." 'INFO'
