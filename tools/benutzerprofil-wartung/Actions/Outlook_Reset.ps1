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
$actionName = 'Outlook_Reset'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log ("Mode={0}" -f $Mode)
Write-Log "=== Outlook Reset gestartet ==="

# Office-Prozessnamen
$officeProcesses = @(
    'outlook','winword','excel','powerpnt','onenote','msaccess','publisher','visio','project',
    'msoia','olk','searchprotocolhost','searchfilterhost'
)

# ==========================
# Office-Prozesse beenden (SDK)
# ==========================
function Stop-OfficeProcesses {
    Write-Log "Beende Office-/Outlook-Prozesse..."
    $killed = WartungsTools.SDK\Stop-SessionProcesses -ProcessNames $officeProcesses -Retries 8 -DelayMs 350
    Write-Log ("Office-Prozesse beendet ({0} gestoppt)." -f $killed)
}

# ==========================
# Cache loeschen (Standard-Reset)
# ==========================
function Clear-OutlookCache {
    Write-Log "Starte Standard-Reset (Cache und temporaere Dateien)..."

    $paths = @(
        (Join-Path $env:APPDATA      'Microsoft\Outlook\RoamCache'),
        (Join-Path $env:APPDATA      'Microsoft\Outlook\Offline Address Books'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCache\Content.Outlook')
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

    Write-Log "Standard-Reset abgeschlossen."
}

# ==========================
# Profil-Reset (inkl. Cache)
# ==========================
function Reset-OutlookProfiles {
    Write-Log "Starte Profil-Reset..."

    Clear-OutlookCache

    $regKeys = @(
        'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles',
        'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover',
        'HKCU:\Software\Microsoft\Office\16.0\Common\Identity'
    )

    foreach ($key in $regKeys) {
        $removed = WartungsTools.SDK\Clear-RegistryPath -Path $key
        if ($removed) { Write-Log ("Registry geloescht: {0}" -f $key) }
        else { Write-Log ("Registry konnte nicht geloescht werden: {0}" -f $key) 'WARN' }
    }

    Write-Log "Profil-Reset abgeschlossen."
}

# ==========================
# Hard-Reset: kompletter Outlook-Schluessel
# ==========================
function Reset-OutlookHard {
    Write-Log "Starte Hard-Reset: kompletter Outlook-Registry-Schluessel..."

    Clear-OutlookCache

    $root = 'HKCU:\Software\Microsoft\Office\16.0\Outlook'
    $removed = WartungsTools.SDK\Clear-RegistryPath -Path $root
    if ($removed) { Write-Log ("Kompletter Outlook-Schluessel geloescht: {0}" -f $root) }
    else { Write-Log ("Outlook-Schluessel konnte nicht geloescht werden: {0}" -f $root) 'WARN' }

    Write-Log "Hard-Reset abgeschlossen."
}

# ==========================
# Office Web Add-ins reparieren (WebView2/Wef)
# ==========================
function Repair-OfficeWebAddins {
    Write-Log "Starte Reparatur: Office-Web-Add-Ins (Wef/WebView2)..."

    $wefWebView2 = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef\webview2'
    if (Test-Path $wefWebView2) {
        $removed = WartungsTools.SDK\Remove-PathSafe -Path $wefWebView2
        if ($removed) { Write-Log ("Web Add-ins bereinigt: {0}" -f $wefWebView2) }
        else { Write-Log ("Web Add-ins nicht vollstaendig bereinigt: {0}" -f $wefWebView2) 'WARN' }
    } else {
        Write-Log ("Wef/WebView2 nicht vorhanden: {0}" -f $wefWebView2)
    }

    Write-Log "Reparatur Office-Web-Add-Ins abgeschlossen."
}

# ==========================
# Hauptfunktion
# ==========================
function Invoke-OutlookReset {
    [CmdletBinding()]
    param(
        [switch]$StandardReset,
        [switch]$ProfileReset,
        [switch]$RepairWebAddins,
        [switch]$HardReset
    )

    Write-Log "Outlook-Reset gestartet..."

    if ($StandardReset -or $ProfileReset -or $HardReset -or $RepairWebAddins) {
        Stop-OfficeProcesses
    }

    if ($HardReset) { Reset-OutlookHard }
    elseif ($ProfileReset) { Reset-OutlookProfiles }
    elseif ($StandardReset) { Clear-OutlookCache }

    if ($RepairWebAddins) { Repair-OfficeWebAddins }

    Write-Log "Outlook-Reset abgeschlossen."
    return @{ ExitCode = 0; Errors = 0; Warnings = 0 }
}

# ==========================
# Ausfuehrung
# ==========================
if ($Mode -eq 'Silent') {
    $standardReset = $true
    $profileReset = $false
    $hardReset = $false
    $repairWebAddins = $false
    if ($Params.ContainsKey('StandardReset')) { $standardReset = [bool]$Params.StandardReset }
    if ($Params.ContainsKey('ProfileReset')) { $profileReset = [bool]$Params.ProfileReset }
    if ($Params.ContainsKey('HardReset')) { $hardReset = [bool]$Params.HardReset }
    if ($Params.ContainsKey('RepairWebAddins')) { $repairWebAddins = [bool]$Params.RepairWebAddins }

    $result = Invoke-OutlookReset -StandardReset:$standardReset -ProfileReset:$profileReset `
        -HardReset:$hardReset -RepairWebAddins:$repairWebAddins
    exit $result.ExitCode
}

# ==========================
# GUI (nur Interactive)
# ==========================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Microsoft Outlook - Reset"; $Form.StartPosition = "CenterScreen"
$Form.Size = New-Object System.Drawing.Size(750, 500); $Form.MaximizeBox = $false

$Panel = New-Object System.Windows.Forms.Panel; $Panel.Dock = 'Fill'; $Panel.AutoScroll = $true
$Form.Controls.Add($Panel)

$script:y = 10
function Add-Label {
    param([string]$Text, [int]$FontSize = 10, [bool]$Bold = $false, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black)
    $label = New-Object System.Windows.Forms.Label
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $(if ($Bold) {[System.Drawing.FontStyle]::Bold} else {[System.Drawing.FontStyle]::Regular}))
    $label.Text = $Text; $label.AutoSize = $true; $label.MaximumSize = New-Object System.Drawing.Size(700, 0)
    $label.Location = New-Object System.Drawing.Point(10, $script:y); $label.ForeColor = $Color
    $Panel.Controls.Add($label); $script:y += $label.Height + 8
}

Add-Label "Microsoft Outlook - Reset" 12 $true
Add-Label "Standard-Reset: loescht den Outlook-Cache und temporaere Dateien (empfohlen)." 10 $false
Add-Label "Profil-Reset: setzt das Outlook-Profil zurueck (inkl. Cache). Profil muss danach neu erstellt werden." 10 $false
Add-Label "Office-Web-Add-Ins reparieren: loescht Wef/WebView2 Cache (Office-Apps werden beendet)." 10 $false
Add-Label "Hard-Reset: Alle Outlook Einstellungen zuruecksetzen (inkl. Cache/Profil)." 10 $false ([System.Drawing.Color]::DarkRed)

$grp = New-Object System.Windows.Forms.GroupBox; $grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $script:y); $grp.Size = New-Object System.Drawing.Size(710, 160)
$Panel.Controls.Add($grp)

$script:innerY = 25
function Add-Check { param([string]$Text); $cb = New-Object System.Windows.Forms.CheckBox; $cb.Text = $Text; $cb.AutoSize = $true; $cb.Location = New-Object System.Drawing.Point(15, $script:innerY); $grp.Controls.Add($cb); $script:innerY += 25; return $cb }

$cbStandard   = Add-Check "Standard-Reset: Cache loeschen (empfohlen)"
$cbProfile    = Add-Check "Profil-Reset: Outlook-Profil zuruecksetzen (inkl. Cache)"
$cbWebAddins  = Add-Check "Office-Web-Add-Ins reparieren (Wef/WebView2 Cache loeschen)"
$cbHard       = Add-Check "Hard-Reset: Alle Outlook Einstellungen zuruecksetzen"

$cbStandard.Checked = $false; $cbProfile.Checked = $false; $cbHard.Checked = $false; $cbWebAddins.Checked = $false

$cbProfile.Add_CheckedChanged({
    if ($cbHard.Checked) { return }
    if ($cbProfile.Checked) { $cbStandard.Checked = $true; $cbStandard.Enabled = $false }
    else { $cbStandard.Enabled = $true }
})

$cbHard.Add_CheckedChanged({
    if ($cbHard.Checked) {
        $cbProfile.Checked = $true; $cbProfile.Enabled = $false
        $cbStandard.Checked = $true; $cbStandard.Enabled = $false
    } else {
        $cbProfile.Enabled = $true; $cbStandard.Enabled = $true
    }
})

$script:y += 180

$btnRun = New-Object System.Windows.Forms.Button; $btnRun.Text = "Reset ausfuehren"; $btnRun.Width = 140
$btnRun.Location = New-Object System.Drawing.Point(20, $script:y); $Panel.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = "Schliessen"; $btnClose.Width = 100
$btnClose.Location = New-Object System.Drawing.Point(180, $script:y); $Panel.Controls.Add($btnClose)
$btnClose.Add_Click({ $Form.Close() }); $Form.AcceptButton = $btnRun; $Form.CancelButton = $btnClose

$btnRun.Add_Click({
    if (-not $cbStandard.Checked -and -not $cbProfile.Checked -and -not $cbHard.Checked -and -not $cbWebAddins.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Aktion auswaehlen.", "Hinweis", 'OK', 'Warning') | Out-Null; return
    }

    if ($cbHard.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG: Beim Hard-Reset wird der komplette Outlook-Registry-Bereich geloescht.`nOutlook verhaelt sich danach wie frisch installiert.`n`nFortfahren?",
            "Hard-Reset bestaetigen", 'YesNo', 'Warning')
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    } elseif ($cbProfile.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show("Outlook-Profil wird geloescht und muss neu eingerichtet werden.`nFortfahren?", "Profil-Reset bestaetigen", 'YesNo', 'Warning')
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    if ($cbWebAddins.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Fuer die Reparatur der Office-Web-Add-Ins muessen alle Office-Anwendungen beendet werden.`nOffene Dokumente bitte vorher speichern.`n`nFortfahren?",
            "Office-Web-Add-Ins reparieren", 'YesNo', 'Warning')
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $btnRun.Enabled = $false
    try {
        Invoke-OutlookReset -StandardReset:$cbStandard.Checked -ProfileReset:$cbProfile.Checked `
            -HardReset:$cbHard.Checked -RepairWebAddins:$cbWebAddins.Checked
        [System.Windows.Forms.MessageBox]::Show("Aktion(en) abgeschlossen.", "Fertig", 'OK', 'Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler:`n$($_.Exception.Message)", "Fehler", 'OK', 'Error') | Out-Null
    } finally { $btnRun.Enabled = $true }
})

[void]$Form.ShowDialog()
Write-Log "=== Outlook Reset beendet ==="
