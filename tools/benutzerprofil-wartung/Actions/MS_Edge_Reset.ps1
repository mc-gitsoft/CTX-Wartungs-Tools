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
$actionName = 'MS_Edge_Reset'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log ("Mode={0}" -f $Mode)
Write-Log "=== Microsoft Edge Reset gestartet ==="

$script:overallSuccess     = $true
$script:favoritesBackupOk  = $false

$script:EdgeUserDataRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data'
$script:EdgeDefaultPath  = Join-Path $script:EdgeUserDataRoot 'Default'
$script:BackupFolder     = Join-Path $env:LOCALAPPDATA 'Wartungsscript\EdgeRestore'

$edgeProcesses = @("msedge", "msedgewebview2")

# ============================================================
# 1) Edge-Prozesse beenden (SDK)
# ============================================================
function Stop-EdgeProcesses {
    Write-Log "Beende Edge-Prozesse des aktuellen Benutzers..."
    $killed = WartungsTools.SDK\Stop-SessionProcesses -ProcessNames $edgeProcesses -Retries 8 -DelayMs 300
    Write-Log ("Edge-Prozesse beendet ({0} gestoppt)." -f $killed)
}

# ============================================================
# 2) Browserdaten loeschen
# ============================================================
function Clear-EdgeBrowserData {
    Write-Log "Starte Loeschung von Browserdaten fuer alle Edge-Profile..."

    if (-not (Test-Path $script:EdgeUserDataRoot)) {
        Write-Log ("Edge User Data Pfad nicht vorhanden: {0}" -f $script:EdgeUserDataRoot) 'WARN'
        return
    }

    $profileDirs = Get-ChildItem -Path $script:EdgeUserDataRoot -Directory -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' -or $_.Name -in @('Guest Profile','System Profile') }

    $cacheSubDirs = @(
        'Cache', 'Code Cache', 'GPUCache', 'Media Cache', 'ShaderCache',
        'Service Worker\CacheStorage', 'Service Worker\ScriptCache', 'Storage\ext'
    )
    $cookieFiles = @('Cookies', 'Cookies-journal', 'Network\Cookies', 'Network\Cookies-journal')
    $siteDirs = @(
        'Local Storage', 'Local Storage\leveldb', 'Session Storage', 'IndexedDB',
        'File System', 'Service Worker', 'Service Worker\CacheStorage',
        'Service Worker\Database', 'Service Worker\ScriptCache', 'Storage'
    )

    foreach ($dir in $profileDirs) {
        foreach ($sub in $cacheSubDirs) {
            $p = Join-Path $dir.FullName $sub
            try {
                if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force -ErrorAction Stop; Write-Log ("Cache geloescht: {0}" -f $p) }
            } catch { Write-Log ("Fehler Cache {0}: {1}" -f $p, $_.Exception.Message) 'WARN'; $script:overallSuccess = $false }
        }
        foreach ($cf in $cookieFiles) {
            $p = Join-Path $dir.FullName $cf
            try {
                if (Test-Path $p) { Remove-Item -Path $p -Force -ErrorAction Stop; Write-Log ("Cookie geloescht: {0}" -f $p) }
            } catch { Write-Log ("Fehler Cookie {0}: {1}" -f $p, $_.Exception.Message) 'WARN'; $script:overallSuccess = $false }
        }
        foreach ($sd in $siteDirs) {
            $p = Join-Path $dir.FullName $sd
            try {
                if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force -ErrorAction Stop; Write-Log ("Site-Daten geloescht: {0}" -f $p) }
            } catch { Write-Log ("Fehler Site-Daten {0}: {1}" -f $p, $_.Exception.Message) 'WARN'; $script:overallSuccess = $false }
        }
    }

    foreach ($g in @('Crashpad', 'SwReporter', 'ShaderCache')) {
        $p = Join-Path $script:EdgeUserDataRoot $g
        try {
            if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force -ErrorAction Stop; Write-Log ("Global-Cache geloescht: {0}" -f $p) }
        } catch { Write-Log ("Fehler Global-Cache {0}: {1}" -f $p, $_.Exception.Message) 'WARN'; $script:overallSuccess = $false }
    }

    Write-Log "Loeschung von Browserdaten abgeschlossen."
}

# ============================================================
# 3) Edge-Registry bereinigen (SDK)
# ============================================================
function Clear-EdgeRegistry {
    Write-Log "Bereinige Edge-Registry (HKCU)..."
    $removed = WartungsTools.SDK\Clear-RegistryPath -Path "HKCU:\Software\Microsoft\Edge"
    if ($removed) { Write-Log "Edge-Registry bereinigt." }
    else { Write-Log "Edge-Registry konnte nicht vollstaendig entfernt werden." 'WARN'; $script:overallSuccess = $false }
}

# ============================================================
# 4) Favoriten sichern / wiederherstellen
# ============================================================
function Backup-EdgeFavorites {
    $script:favoritesBackupOk = $false
    $favoritesFile = Join-Path $script:EdgeDefaultPath 'Bookmarks'
    $firstRunFile  = Join-Path $script:EdgeUserDataRoot 'First Run'

    if (-not (Test-Path $favoritesFile) -and -not (Test-Path $firstRunFile)) { return }

    if (-not (Test-Path $script:BackupFolder)) {
        try { New-Item -ItemType Directory -Path $script:BackupFolder -ErrorAction Stop | Out-Null }
        catch { Write-Log ("Backup-Ordner Fehler: {0}" -f $_.Exception.Message) 'ERROR'; $script:overallSuccess = $false; return }
    }

    if (Test-Path $favoritesFile) {
        try { Copy-Item -Path $favoritesFile -Destination $script:BackupFolder -Force; $script:favoritesBackupOk = $true; Write-Log "Favoriten gesichert." }
        catch { Write-Log ("Favoriten-Backup Fehler: {0}" -f $_.Exception.Message) 'ERROR'; $script:overallSuccess = $false }
    }
    if (Test-Path $firstRunFile) {
        try { Copy-Item -Path $firstRunFile -Destination (Join-Path $script:BackupFolder 'First Run') -Force; Write-Log "'First Run' gesichert." }
        catch { Write-Log ("First-Run-Backup Fehler: {0}" -f $_.Exception.Message) 'WARN' }
    }
}

function Restore-EdgeFavorites {
    if (-not (Test-Path $script:EdgeUserDataRoot)) {
        try { New-Item -ItemType Directory -Path $script:EdgeUserDataRoot -ErrorAction Stop | Out-Null }
        catch { Write-Log ("Edge User Data Root Fehler: {0}" -f $_.Exception.Message) 'ERROR'; $script:overallSuccess = $false; return }
    }

    $backupFirstRun = Join-Path $script:BackupFolder 'First Run'
    if (Test-Path $backupFirstRun) {
        try { Copy-Item -Path $backupFirstRun -Destination (Join-Path $script:EdgeUserDataRoot 'First Run') -Force; Write-Log "'First Run' wiederhergestellt." }
        catch { Write-Log ("First-Run Restore Fehler: {0}" -f $_.Exception.Message) 'WARN' }
    }

    if (-not $script:favoritesBackupOk) { Write-Log "Kein Favoriten-Backup - ueberspringe Wiederherstellung." 'WARN'; return }

    $backupFavoritesFile = Join-Path $script:BackupFolder 'Bookmarks'
    if (-not (Test-Path $backupFavoritesFile)) { return }

    try {
        if (-not (Test-Path $script:EdgeDefaultPath)) { New-Item -ItemType Directory -Path $script:EdgeDefaultPath -ErrorAction Stop | Out-Null }
        Copy-Item -Path $backupFavoritesFile -Destination $script:EdgeDefaultPath -Force
        Write-Log "Favoriten wiederhergestellt."
    } catch { Write-Log ("Favoriten Restore Fehler: {0}" -f $_.Exception.Message) 'ERROR'; $script:overallSuccess = $false }
}

function Cleanup-EdgeBackup {
    if (Test-Path $script:BackupFolder) {
        try { Remove-Item -Path $script:BackupFolder -Recurse -Force -ErrorAction Stop } catch { }
    }
}

# ============================================================
# 5) Hard-Reset
# ============================================================
function Reset-EdgeHard {
    Write-Log "Starte Hard-Reset fuer Edge..."
    Backup-EdgeFavorites

    foreach ($p in @((Join-Path $env:LOCALAPPDATA 'Microsoft\Edge'), (Join-Path $env:APPDATA 'Microsoft\Edge'))) {
        if (Test-Path $p) {
            $removed = WartungsTools.SDK\Remove-PathSafe -Path $p
            if ($removed) { Write-Log ("Edge-Ordner geloescht: {0}" -f $p) }
            else { Write-Log ("Edge-Ordner nicht vollstaendig geloescht: {0}" -f $p) 'WARN'; $script:overallSuccess = $false }
        }
    }

    Clear-EdgeRegistry

    $script:EdgeUserDataRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data'
    $script:EdgeDefaultPath  = Join-Path $script:EdgeUserDataRoot 'Default'

    Restore-EdgeFavorites
    Cleanup-EdgeBackup
    Write-Log "Hard-Reset abgeschlossen."
}

# ============================================================
# 6) Hauptfunktion
# ============================================================
function Invoke-EdgeReset {
    [CmdletBinding()]
    param([switch]$BrowserDataReset, [switch]$HardReset)

    Write-Log "Edge-Reset gestartet..."
    Stop-EdgeProcesses
    if ($HardReset) { Reset-EdgeHard }
    elseif ($BrowserDataReset) { Clear-EdgeBrowserData }
    Write-Log "Edge-Reset abgeschlossen."

    return @{ ExitCode = $(if ($script:overallSuccess) { 0 } else { 1 }); Errors = 0; Warnings = 0 }
}

# ============================================================
# Ausfuehrung
# ============================================================
if ($Mode -eq 'Silent') {
    $script:overallSuccess = $true; $script:favoritesBackupOk = $false
    $browserDataReset = $true; $hardReset = $false
    if ($Params.ContainsKey('HardReset')) { $hardReset = [bool]$Params.HardReset }
    if ($Params.ContainsKey('BrowserDataReset')) { $browserDataReset = [bool]$Params.BrowserDataReset }
    if ($hardReset) { $browserDataReset = $false }

    $result = Invoke-EdgeReset -BrowserDataReset:$browserDataReset -HardReset:$hardReset
    exit $result.ExitCode
}

# ============================================================
# GUI (nur Interactive)
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Microsoft Edge - Reset"; $Form.StartPosition = "CenterScreen"
$Form.Size = New-Object System.Drawing.Size(780, 460); $Form.MaximizeBox = $false

$Panel = New-Object System.Windows.Forms.Panel; $Panel.Dock = 'Fill'; $Panel.AutoScroll = $true
$Form.Controls.Add($Panel)

$script:y = 10
function Add-Label {
    param([string]$Text, [int]$FontSize = 10, [bool]$Bold = $false, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black)
    $label = New-Object System.Windows.Forms.Label
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $(if ($Bold) {[System.Drawing.FontStyle]::Bold} else {[System.Drawing.FontStyle]::Regular}))
    $label.Text = $Text; $label.AutoSize = $true; $label.MaximumSize = New-Object System.Drawing.Size(730, 0)
    $label.Location = New-Object System.Drawing.Point(10, $script:y); $label.ForeColor = $Color
    $Panel.Controls.Add($label); $script:y += $label.Height + 8
}

Add-Label "Microsoft Edge - Reset" 12 $true
Add-Label "Browserdaten loeschen: entfernt Cache, Cookies und Website-Daten. Favoriten bleiben erhalten (empfohlen)." 10 $false ([System.Drawing.Color]::DarkRed)
Add-Label "Hard-Reset: setzt Edge komplett zurueck. Favoriten und 'First Run' werden gesichert." 10 $false ([System.Drawing.Color]::DarkRed)

$grp = New-Object System.Windows.Forms.GroupBox; $grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $script:y); $grp.Size = New-Object System.Drawing.Size(740, 130)
$Panel.Controls.Add($grp)

$script:innerY = 25
function Add-Check { param([string]$Text); $cb = New-Object System.Windows.Forms.CheckBox; $cb.Text = $Text; $cb.AutoSize = $true; $cb.Location = New-Object System.Drawing.Point(15, $script:innerY); $grp.Controls.Add($cb); $script:innerY += 25; return $cb }

$cbBrowser = Add-Check "Browserdaten loeschen (Cache + Cookies + Website-Daten) - empfohlen"
$cbHard    = Add-Check "Hard-Reset: Edge komplett zuruecksetzen (AppData + Registry, Favoriten bleiben)"
$cbBrowser.Checked = $true; $cbHard.Checked = $false

$cbHard.Add_CheckedChanged({ if ($cbHard.Checked) { $cbBrowser.Checked = $false; $cbBrowser.Enabled = $false } else { $cbBrowser.Enabled = $true; $cbBrowser.Checked = $true } })

$script:y += 150
$btnRun = New-Object System.Windows.Forms.Button; $btnRun.Text = "Reset ausfuehren"; $btnRun.Width = 140; $btnRun.Location = New-Object System.Drawing.Point(20, $script:y); $Panel.Controls.Add($btnRun)
$btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = "Schliessen"; $btnClose.Width = 100; $btnClose.Location = New-Object System.Drawing.Point(180, $script:y); $Panel.Controls.Add($btnClose)
$btnClose.Add_Click({ $Form.Close() }); $Form.AcceptButton = $btnRun; $Form.CancelButton = $btnClose

$btnRun.Add_Click({
    if (-not $cbBrowser.Checked -and -not $cbHard.Checked) { [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Aktion auswaehlen.", "Hinweis", 'OK', 'Warning') | Out-Null; return }
    if ($cbHard.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show("ACHTUNG: Hard-Reset entfernt alle Edge-Daten.`nFavoriten werden gesichert.`n`nFortfahren?", "Hard-Reset bestaetigen", 'YesNo', 'Warning')
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    } elseif ($cbBrowser.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show("Browserdaten werden geloescht. Sie werden abgemeldet.`n`nFortfahren?", "Bestaetigen", 'YesNo', 'Warning')
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    $btnRun.Enabled = $false; $script:overallSuccess = $true; $script:favoritesBackupOk = $false
    try {
        $null = Invoke-EdgeReset -BrowserDataReset:$cbBrowser.Checked -HardReset:$cbHard.Checked
        [System.Windows.Forms.MessageBox]::Show("Reset abgeschlossen.", "Fertig", 'OK', 'Information') | Out-Null
    } catch { [System.Windows.Forms.MessageBox]::Show("Fehler:`n$($_.Exception.Message)", "Fehler", 'OK', 'Error') | Out-Null }
    finally { $btnRun.Enabled = $true }
})

[void]$Form.ShowDialog()
Write-Log "=== Microsoft Edge Reset beendet ==="
