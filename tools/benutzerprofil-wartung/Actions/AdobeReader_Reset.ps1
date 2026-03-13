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
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log ("Mode={0}" -f $Mode)
Write-Log "=== Adobe Acrobat Reset gestartet ==="

# Prozessnamen
$acrobatProcesses = @(
    'Acrobat','AcroCEF','AcroBroker','AdobeCollabSync','AdobeIPCBroker'
)

# ==========================
# Hauptfunktion
# ==========================
function Invoke-AdobeReset {
    [CmdletBinding()]
    param()

    Write-Log "Starte Adobe Acrobat Reset..."

    # 1) Prozesse beenden (SDK)
    Write-Log "Beende Acrobat-Prozesse des aktuellen Benutzers..."
    $killed = WartungsTools.SDK\Stop-SessionProcesses -ProcessNames $acrobatProcesses -Retries 5 -DelayMs 400
    Write-Log ("Acrobat-Prozesse beendet ({0} gestoppt)." -f $killed)

    # 2) Benutzerpfade bereinigen (SDK)
    Write-Log "Loesche Acrobat-Ordner im Benutzerprofil..."

    $paths = @(
        (Join-Path $env:LOCALAPPDATA 'Adobe\Acrobat'),
        (Join-Path $env:APPDATA      'Adobe\Acrobat')
    )

    foreach ($p in $paths) {
        if (Test-Path $p) {
            $removed = WartungsTools.SDK\Remove-PathSafe -Path $p
            if ($removed) { Write-Log ("Geloescht: {0}" -f $p) }
            else { Write-Log ("Konnte nicht vollstaendig geloescht werden: {0}" -f $p) 'WARN' }
        } else {
            Write-Log ("Nicht vorhanden: {0}" -f $p)
        }
    }

    # 3) Registry bereinigen (SDK)
    Write-Log "Entferne Acrobat-Registry-Eintraege (HKCU)..."

    $regKey = 'HKCU:\Software\Adobe\Adobe Acrobat'
    $removed = WartungsTools.SDK\Clear-RegistryPath -Path $regKey
    if ($removed) { Write-Log ("Registry geloescht: {0}" -f $regKey) }
    else { Write-Log ("Registry konnte nicht geloescht werden: {0}" -f $regKey) 'WARN' }

    Write-Log "Adobe Acrobat Reset abgeschlossen."
    return @{ ExitCode = 0; Errors = 0; Warnings = 0 }
}

# ==========================
# Ausfuehrung
# ==========================
if ($Mode -eq 'Silent') {
    $result = Invoke-AdobeReset
    exit $result.ExitCode
}

# ==========================
# Interactive: einfache Ausfuehrung mit Benutzerausgabe
# ==========================
$result = Invoke-AdobeReset

Write-Host ""
Write-Host "Adobe Acrobat wurde im Benutzerprofil so weit wie moeglich zurueckgesetzt." -ForegroundColor Green
Write-Host "Beim naechsten Start werden Standarddaten und Konfigurationen neu angelegt."
Write-Host ""

Write-Log "=== Adobe Acrobat Reset beendet ==="
