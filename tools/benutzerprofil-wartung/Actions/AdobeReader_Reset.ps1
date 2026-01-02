[CmdletBinding()]
param(
    [ValidateSet('Interactive','Silent')]
    [string]$Mode = 'Interactive',
    [object]$Params = @{}
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot 'shared\WartungsTools.SDK.psm1') -Force
$Params = ConvertTo-Hashtable -InputObject $Params
$toolId = (Get-Content (Join-Path $toolRoot 'tool.json') -Raw | ConvertFrom-Json).toolId
$actionName = 'AdobeReader_Reset'

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log "=== Adobe Acrobat Reset gestartet ==="
# =====================================================================
# 1) Helper: Acrobat-Prozesse nur in der aktuellen User-Session beenden
# =====================================================================
function Stop-AcrobatProcessesUserScoped {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ProcessNames,

        [int]$Retries = 5,
        [int]$DelayMs = 400
    )

    try {
        $currentSession = (Get-Process -Id $PID).SessionId
        Write-Log "Aktuelle SessionId: $currentSession"
    }
    catch {
        Write-Log "Konnte aktuelle SessionId nicht ermitteln, beende Prozesse ohne Session-Filter." 'WARN'
        $currentSession = $null
    }

    for ($i = 1; $i -le $Retries; $i++) {
        try {
            if ($currentSession -ne $null) {
                $candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                    $_.Name -in $ProcessNames -and $_.SessionId -eq $currentSession
                }
            }
            else {
                $candidates = Get-Process -Name $ProcessNames -ErrorAction SilentlyContinue
            }

            if (-not $candidates) {
                if ($i -eq 1) {
                    Write-Log "Keine Acrobat-Prozesse in der aktuellen Session gefunden."
                }
            }
            else {
                $names = ($candidates | Select-Object -ExpandProperty Name -Unique) -join ', '
                Write-Log ("Versuch {0}/{1}: Beende Prozesse: {2}" -f $i, $Retries, $names)

                foreach ($p in $candidates) {
                    try {
                        Stop-Process -Id $p.Id -Force -ErrorAction Stop
                        Write-Log ("Prozess beendet: {0} (PID {1})" -f $p.Name, $p.Id)
                    }
                    catch {
                        Write-Log ("Fehler beim Beenden von {0} (PID {1}): {2}" -f $p.Name, $p.Id, $_.Exception.Message) 'WARN'
                    }
                }
            }
        }
        catch {
            Write-Log "Fehler beim Beenden von Acrobat-Prozessen: $($_.Exception.Message)" 'WARN'
        }

        Start-Sleep -Milliseconds $DelayMs
    }
}

# =====================================================================
# 2) Prozesse im Benutzerkontext beenden
# =====================================================================
$acrobatProcesses = @(
    'Acrobat',
    'AcroCEF',
    'AcroBroker',
    'AdobeCollabSync',
    'AdobeIPCBroker'
)

Write-Log "[1/3] Beende Acrobat-Prozesse des aktuellen Benutzers..." 'INFO'
Stop-AcrobatProcessesUserScoped -ProcessNames $acrobatProcesses

# =====================================================================
# 3) Helper-Funktionen für sichere Löschung von Ordnern / Registry
# =====================================================================
function Remove-FolderSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log "Ordner nicht vorhanden: $Path"
            return
        }

        Write-Log "Lösche Ordner: $Path"
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        Write-Log "Ordner gelöscht: $Path"
    }
    catch {
        Write-Log ("Fehler beim Löschen von Ordner '{0}': {1}" -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Remove-KeySafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Key
    )

    try {
        if (-not (Test-Path $Key)) {
            Write-Log "Registry-Zweig nicht vorhanden: $Key"
            return
        }

        Write-Log "Lösche Registry-Zweig: $Key"
        Remove-Item -Path $Key -Recurse -Force -ErrorAction Stop
        Write-Log "Registry-Zweig gelöscht: $Key"
    }
    catch {
        Write-Log ("Fehler beim Löschen des Registry-Zweigs '{0}': {1}" -f $Key, $_.Exception.Message) 'WARN'
    }
}

# =====================================================================
# 4) Benutzerpfade & Registry von Adobe Acrobat bereinigen
# =====================================================================
Write-Log "[2/3] Lösche Acrobat-Ordner im Benutzerprofil..." 'INFO'

$LocalAcrobat   = Join-Path $env:LOCALAPPDATA 'Adobe\Acrobat'
$RoamingAcrobat = Join-Path $env:APPDATA      'Adobe\Acrobat'

Remove-FolderSafe -Path $LocalAcrobat
Remove-FolderSafe -Path $RoamingAcrobat

Write-Log "[3/3] Entferne Acrobat-Registry-Einträge (HKCU)..." 'INFO'

$RegAcrobat = 'HKCU:\Software\Adobe\Adobe Acrobat'
Remove-KeySafe -Key $RegAcrobat

Write-Log "Adobe Acrobat Reset im Benutzerprofil abgeschlossen." 'INFO'

Write-Host ""
Write-Host "Adobe Acrobat wurde im Benutzerprofil so weit wie möglich zurückgesetzt." -ForegroundColor Green
Write-Host "Beim nächsten Start werden Standarddaten und Konfigurationen neu angelegt."
Write-Host ""


