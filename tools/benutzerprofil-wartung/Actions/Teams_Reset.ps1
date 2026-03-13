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
$actionName = 'Teams_Reset'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log ("Mode={0}" -f $Mode)

# ==========================
# Prozessnamen
# ==========================
$teamsProcesses = @(
    'ms-teams','teams','teamswebview2','msedgewebview2',
    'teams-updater','teamsbootstrapper','update','teamsupdate',
    'microsoft.aad.brokerplugin','oneauth','tokenbroker'
)

# ==========================
# Sichere Pfad-Loeschung (nutzt SDK Remove-PathSafe)
# ==========================
function Remove-TeamsPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Nicht vorhanden: {0}" -f $Path)
        return
    }

    $result = WartungsTools.SDK\Remove-PathSafe -Path $Path
    if ($result) {
        Write-Log ("Geloescht: {0}" -f $Path)
    } else {
        Write-Log ("Konnte nicht vollstaendig geloescht werden: {0}" -f $Path) 'WARN'
    }
}

# ==========================
# Hard-Reset: Teams-AppX fuer Benutzer entfernen
# ==========================
function Invoke-TeamsAppxReset {
    Write-Log "Hard-Reset: Entferne Teams-AppX (MSTeams) fuer diesen Benutzer..."

    try {
        $pkg = Get-AppxPackage -Name 'MSTeams' -ErrorAction Stop
        Write-Log ("Gefundenes AppX-Paket: {0}" -f $pkg.PackageFullName)
    } catch {
        Write-Log "Teams-AppX fuer diesen Benutzer nicht gefunden - Hard-Reset uebersprungen." 'WARN'
        return
    }

    try {
        Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop
        Write-Log "Teams-AppX erfolgreich entfernt. Teams wird beim naechsten Login automatisch neu bereitgestellt."
    } catch {
        Write-Log ("Remove-AppxPackage Fehler: {0}" -f $_.Exception.Message) 'ERROR'
    }

    # Classic-Teams-Dateipfade (falls vorhanden) aufraeumen
    $classicDirs = @(
        (Join-Path $env:APPDATA      'Microsoft\Teams'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Teams')
    )
    foreach ($d in $classicDirs) {
        Remove-TeamsPath -Path $d
    }

    # Classic-Teams-Registry (falls vorhanden) entfernen
    $regKeys = @(
        'HKCU:\Software\Microsoft\Office\Teams',
        'HKCU:\Software\Microsoft\Teams'
    )
    foreach ($r in $regKeys) {
        $removed = WartungsTools.SDK\Clear-RegistryPath -Path $r
        if ($removed) {
            Write-Log ("Registry geloescht: {0}" -f $r)
        }
    }

    Write-Log "Hard-Reset abgeschlossen. Benutzer muss sich neu anmelden, damit Teams neu bereitgestellt wird."
}

# ==========================
# Hauptfunktion
# ==========================
function Invoke-Teams2Reset {
    [CmdletBinding()]
    param(
        [switch]$ClearCache,
        [switch]$HardReset
    )

    Write-Log "Starte Teams-Reset..."

    # 1. Prozesse beenden (SDK-Hilfsfunktion mit taskkill-Fallback)
    Write-Log "Beende Teams-Prozesse..."
    $killed = WartungsTools.SDK\Stop-SessionProcesses -ProcessNames $teamsProcesses -Retries 5 -DelayMs 400 -UseTaskkillFallback
    Write-Log ("Prozesse beendet: {0}" -f $killed)

    # 2. Hard-Reset?
    if ($HardReset) {
        Invoke-TeamsAppxReset
        Write-Log "Hard-Reset ausgefuehrt."
        return @{ ExitCode = 0; Errors = 0; Warnings = 0 }
    }

    # 3. Normaler Reset: nur Cache loeschen
    if ($ClearCache) {
        $pkgBase = Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe'
        $targets = @(
            (Join-Path $pkgBase 'LocalCache\Microsoft\MSTeams'),
            (Join-Path $pkgBase 'TempState')
        )

        foreach ($t in $targets) {
            Remove-TeamsPath -Path $t
        }
        Write-Log "Teams-Cache-Reset abgeschlossen."
    } else {
        Write-Log "Kein Cache-Reset ausgewaehlt."
    }

    return @{ ExitCode = 0; Errors = 0; Warnings = 0 }
}

# ==========================
# Ausfuehrung
# ==========================
if ($Mode -eq 'Silent') {
    # Silent-Modus: Parameter aus Params auswerten
    $clearCache = $true
    $hardReset = $false
    if ($Params.ContainsKey('HardReset')) { $hardReset = [bool]$Params.HardReset }
    if ($Params.ContainsKey('ClearCache')) { $clearCache = [bool]$Params.ClearCache }
    if ($hardReset) { $clearCache = $false }

    $result = Invoke-Teams2Reset -ClearCache:$clearCache -HardReset:$hardReset
    exit $result.ExitCode
}

# ==========================
# GUI (nur Interactive)
# ==========================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form               = New-Object System.Windows.Forms.Form
$Form.Text          = "Microsoft Teams 2.0 - Reset"
$Form.StartPosition = "CenterScreen"
$Form.Size          = New-Object System.Drawing.Size(700, 420)
$Form.MaximizeBox   = $false

$Panel = New-Object System.Windows.Forms.Panel
$Panel.Dock = 'Fill'
$Panel.AutoScroll = $true
$Form.Controls.Add($Panel)

$script:y = 10

function Add-Label {
    param(
        [string]$Text,
        [int]$FontSize = 10,
        [bool]$Bold = $false,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )
    $label = New-Object System.Windows.Forms.Label
    $style = [System.Drawing.FontStyle]::Regular
    if ($Bold) { $style = [System.Drawing.FontStyle]::Bold }
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $style)
    $label.Text = $Text
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(660, 0)
    $label.Location = New-Object System.Drawing.Point(10, $script:y)
    $label.ForeColor = $Color
    $Panel.Controls.Add($label)
    $script:y += $label.Height + 8
}

Add-Label "Microsoft Teams 2.0 - Reset" 12 $true
Add-Label "Empfohlen: Teams-Cache loeschen (schneller Standard-Reset bei Problemen)." 10 $false ([System.Drawing.Color]::DarkRed)
Add-Label "Hard-Reset: entfernt Teams vollstaendig fuer diesen Benutzer. Eine Neuanmeldung ist erforderlich, danach wird Teams neu bereitgestellt. Nur verwenden, wenn der Cache-Reset nicht hilft." 9 $false

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $script:y)
$grp.Size = New-Object System.Drawing.Size(660, 120)
$Panel.Controls.Add($grp)

$script:innerY = 25
function Add-Check {
    param([string]$Text)
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $Text
    $cb.AutoSize = $true
    $cb.Location = New-Object System.Drawing.Point(15, $script:innerY)
    $grp.Controls.Add($cb)
    $script:innerY += 25
    return $cb
}

$cbCache = Add-Check "Teams-Cache loeschen (empfohlen)"
$cbHard  = Add-Check "Hard-Reset: Teams-App entfernen (Neuanmeldung notwendig)"

$cbCache.Checked = $true
$cbHard.Checked  = $false

$cbHard.Add_CheckedChanged({
    if ($cbHard.Checked) {
        $cbCache.Checked = $false
        $cbCache.Enabled = $false
    } else {
        $cbCache.Enabled = $true
        $cbCache.Checked = $true
    }
})

$script:y += 140

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Reset ausfuehren"
$btnRun.Width = 150
$btnRun.Location = New-Object System.Drawing.Point(20, $script:y)
$Panel.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Schliessen"
$btnClose.Width = 100
$btnClose.Location = New-Object System.Drawing.Point(190, $script:y)
$Panel.Controls.Add($btnClose)

$btnClose.Add_Click({ $Form.Close() })

$Form.AcceptButton = $btnRun
$Form.CancelButton = $btnClose

$btnRun.Add_Click({
    if (-not $cbCache.Checked -and -not $cbHard.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte waehlen Sie mindestens eine Aktion aus (Cache-Reset oder Hard-Reset).",
            "Keine Aktion ausgewaehlt",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $btnRun.Enabled = $false
    try {
        $result = Invoke-Teams2Reset `
            -ClearCache:$cbCache.Checked `
            -HardReset:$cbHard.Checked

        [System.Windows.Forms.MessageBox]::Show(
            "Reset abgeschlossen.",
            "Fertig",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler beim Ausfuehren des Resets:`n$($_.Exception.Message)",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $btnRun.Enabled = $true
    }
})

[void]$Form.ShowDialog()
