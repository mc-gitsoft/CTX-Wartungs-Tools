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
$actionName = 'Google_Chrome_Reset'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

Write-Log ("Mode={0}" -f $Mode)
Write-Log "=== Google Chrome Reset gestartet ==="

# globale Statusvariablen
$script:overallSuccess   = $true
$script:bookmarkBackupOk = $false

# Standardpfade
$script:ChromeUserDataRoot   = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
$script:ChromeDefaultPath    = Join-Path $script:ChromeUserDataRoot 'Default'
$script:BackupFolder         = Join-Path $env:LOCALAPPDATA 'Wartungsscript\ChromeRestore'

# Prozessnamen
$chromeProcesses = @("chrome", "googleupdate", "googlecrashhandler")

# ============================================================
# 1) Chrome-Prozesse beenden (SDK)
# ============================================================
function Stop-ChromeProcesses {
    Write-Log "Beende Chrome-Prozesse des aktuellen Benutzers..."
    $killed = WartungsTools.SDK\Stop-SessionProcesses -ProcessNames $chromeProcesses -Retries 8 -DelayMs 300
    Write-Log ("Chrome-Prozesse beendet ({0} gestoppt)." -f $killed)
}

# ============================================================
# 2) Browserdaten loeschen (Cache + Cookies + Website-Daten, alle Profile)
# ============================================================
function Clear-ChromeBrowserData {
    Write-Log "Starte Loeschung von Browserdaten (Cache + Cookies + Website-Daten) fuer alle Chrome-Profile..."

    if (-not (Test-Path $script:ChromeUserDataRoot)) {
        Write-Log ("Chrome User Data Pfad nicht vorhanden: {0}" -f $script:ChromeUserDataRoot) 'WARN'
        return
    }

    $profileDirs = Get-ChildItem -Path $script:ChromeUserDataRoot -Directory -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' -or $_.Name -in @('Guest Profile','System Profile') }

    $cacheSubDirs = @(
        'Cache', 'Code Cache', 'GPUCache', 'Media Cache', 'ShaderCache',
        'Service Worker\CacheStorage', 'Service Worker\ScriptCache', 'Storage\ext'
    )

    $cookieFiles = @(
        'Cookies', 'Cookies-journal', 'Network\Cookies', 'Network\Cookies-journal'
    )

    $siteDirs = @(
        'Local Storage', 'Local Storage\leveldb', 'Session Storage', 'IndexedDB',
        'File System', 'Service Worker', 'Service Worker\CacheStorage',
        'Service Worker\Database', 'Service Worker\ScriptCache', 'Storage'
    )

    foreach ($dir in $profileDirs) {
        foreach ($sub in $cacheSubDirs) {
            $p = Join-Path $dir.FullName $sub
            try {
                if (Test-Path $p) {
                    Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                    Write-Log ("Cache-Ordner geloescht: {0}" -f $p)
                }
            } catch {
                Write-Log ("Fehler beim Loeschen von Cache-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }

        foreach ($cf in $cookieFiles) {
            $p = Join-Path $dir.FullName $cf
            try {
                if (Test-Path $p) {
                    Remove-Item -Path $p -Force -ErrorAction Stop
                    Write-Log ("Cookies-Datei geloescht: {0}" -f $p)
                }
            } catch {
                Write-Log ("Fehler beim Loeschen von Cookies-Datei {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }

        foreach ($sd in $siteDirs) {
            $p = Join-Path $dir.FullName $sd
            try {
                if (Test-Path $p) {
                    Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                    Write-Log ("Site-Daten-Ordner geloescht: {0}" -f $p)
                }
            } catch {
                Write-Log ("Fehler beim Loeschen von Site-Daten-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }
    }

    $globalCacheDirs = @('Crashpad', 'SwReporter', 'ShaderCache')
    foreach ($g in $globalCacheDirs) {
        $p = Join-Path $script:ChromeUserDataRoot $g
        try {
            if (Test-Path $p) {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Log ("Globaler Cache-Ordner geloescht: {0}" -f $p)
            }
        } catch {
            Write-Log ("Fehler beim Loeschen von globalem Cache-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }

    Write-Log "Loeschung von Browserdaten abgeschlossen."
}

# ============================================================
# 3) Chrome-Registry (HKCU) bereinigen (SDK)
# ============================================================
function Clear-ChromeRegistry {
    Write-Log "Bereinige Chrome-Registry-Eintraege (HKCU)..."

    $chromeRootKey = "HKCU:\Software\Google\Chrome"
    $removed = WartungsTools.SDK\Clear-RegistryPath -Path $chromeRootKey
    if ($removed) {
        Write-Log ("Chrome-Registry bereinigt: {0}" -f $chromeRootKey)
    } else {
        Write-Log ("Chrome-Registry konnte nicht vollstaendig entfernt werden (Policy/Berechtigung): {0}" -f $chromeRootKey) 'WARN'
        $script:overallSuccess = $false
    }
}

# ============================================================
# 4) Bookmarks + "First Run" sichern / wiederherstellen
# ============================================================
function Backup-ChromeBookmarks {
    $script:bookmarkBackupOk = $false

    $bookmarksFile = Join-Path $script:ChromeDefaultPath 'Bookmarks'
    $firstRunFile  = Join-Path $script:ChromeUserDataRoot 'First Run'

    if (-not (Test-Path $bookmarksFile) -and -not (Test-Path $firstRunFile)) {
        Write-Log "Keine Bookmarks oder First-Run-Datei zum Sichern gefunden."
        return
    }

    if (-not (Test-Path $script:BackupFolder)) {
        try {
            New-Item -ItemType Directory -Path $script:BackupFolder -ErrorAction Stop | Out-Null
            Write-Log ("Backup-Ordner erstellt: {0}" -f $script:BackupFolder)
        } catch {
            Write-Log ("Fehler beim Erstellen des Backup-Ordners: {0}" -f $_.Exception.Message) 'ERROR'
            $script:overallSuccess = $false
            return
        }
    }

    if (Test-Path $bookmarksFile) {
        try {
            Copy-Item -Path $bookmarksFile -Destination $script:BackupFolder -Force
            Write-Log ("Lesezeichen gesichert nach: {0}" -f $script:BackupFolder)
            $script:bookmarkBackupOk = $true
        } catch {
            Write-Log ("Fehler beim Sichern der Lesezeichen: {0}" -f $_.Exception.Message) 'ERROR'
            $script:overallSuccess   = $false
            $script:bookmarkBackupOk = $false
        }
    }

    if (Test-Path $firstRunFile) {
        try {
            Copy-Item -Path $firstRunFile -Destination (Join-Path $script:BackupFolder 'First Run') -Force
            Write-Log ("'First Run'-Datei gesichert: {0}" -f $firstRunFile)
        } catch {
            Write-Log ("Fehler beim Sichern der 'First Run'-Datei: {0}" -f $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }
}

function Restore-ChromeBookmarks {
    if (-not (Test-Path $script:ChromeUserDataRoot)) {
        try {
            New-Item -ItemType Directory -Path $script:ChromeUserDataRoot -ErrorAction Stop | Out-Null
        } catch {
            Write-Log ("Fehler beim Erstellen von Chrome User Data Root: {0}" -f $_.Exception.Message) 'ERROR'
            $script:overallSuccess = $false
            return
        }
    }

    $backupFirstRun = Join-Path $script:BackupFolder 'First Run'
    if (Test-Path $backupFirstRun) {
        try {
            Copy-Item -Path $backupFirstRun -Destination (Join-Path $script:ChromeUserDataRoot 'First Run') -Force
            Write-Log "'First Run'-Datei wiederhergestellt."
        } catch {
            Write-Log ("Fehler beim Wiederherstellen der 'First Run'-Datei: {0}" -f $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }

    if (-not $script:bookmarkBackupOk) {
        Write-Log "Bookmark-Backup war nicht erfolgreich - keine Wiederherstellung." 'WARN'
        return
    }

    $backupBookmarksFile = Join-Path $script:BackupFolder 'Bookmarks'
    if (-not (Test-Path $backupBookmarksFile)) { return }

    try {
        if (-not (Test-Path $script:ChromeDefaultPath)) {
            New-Item -ItemType Directory -Path $script:ChromeDefaultPath -ErrorAction Stop | Out-Null
        }
        Copy-Item -Path $backupBookmarksFile -Destination $script:ChromeDefaultPath -Force
        Write-Log "Lesezeichen erfolgreich wiederhergestellt."
    } catch {
        Write-Log ("Fehler beim Wiederherstellen der Lesezeichen: {0}" -f $_.Exception.Message) 'ERROR'
        $script:overallSuccess = $false
    }
}

function Cleanup-BookmarkBackup {
    if (Test-Path $script:BackupFolder) {
        try {
            Remove-Item -Path $script:BackupFolder -Recurse -Force -ErrorAction Stop
        } catch {
            Write-Log ("Fehler beim Loeschen des Bookmark-Backups: {0}" -f $_.Exception.Message) 'WARN'
        }
    }
}

# ============================================================
# 5) Hard-Reset
# ============================================================
function Reset-ChromeHard {
    Write-Log "Starte Hard-Reset fuer Chrome..."

    Backup-ChromeBookmarks

    $localChrome   = Join-Path $env:LOCALAPPDATA 'Google\Chrome'
    $roamingChrome = Join-Path $env:APPDATA      'Google\Chrome'

    foreach ($p in @($localChrome, $roamingChrome)) {
        if (Test-Path $p) {
            $removed = WartungsTools.SDK\Remove-PathSafe -Path $p
            if ($removed) {
                Write-Log ("Chrome-Ordner geloescht: {0}" -f $p)
            } else {
                Write-Log ("Chrome-Ordner konnte nicht vollstaendig geloescht werden: {0}" -f $p) 'WARN'
                $script:overallSuccess = $false
            }
        }
    }

    Clear-ChromeRegistry

    $script:ChromeUserDataRoot = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
    $script:ChromeDefaultPath  = Join-Path $script:ChromeUserDataRoot 'Default'

    Restore-ChromeBookmarks
    Cleanup-BookmarkBackup

    Write-Log "Hard-Reset abgeschlossen."
}

# ============================================================
# 6) Hauptfunktion
# ============================================================
function Invoke-ChromeReset {
    [CmdletBinding()]
    param(
        [switch]$BrowserDataReset,
        [switch]$HardReset
    )

    Write-Log "Chrome-Reset gestartet..."

    Stop-ChromeProcesses

    if ($HardReset) {
        Reset-ChromeHard
    } elseif ($BrowserDataReset) {
        Clear-ChromeBrowserData
    }

    Write-Log "Chrome-Reset abgeschlossen."

    return @{
        ExitCode = $(if ($script:overallSuccess) { 0 } else { 1 })
        Errors   = 0
        Warnings = 0
    }
}

# ============================================================
# Ausfuehrung
# ============================================================
if ($Mode -eq 'Silent') {
    $script:overallSuccess   = $true
    $script:bookmarkBackupOk = $false

    $browserDataReset = $true
    $hardReset = $false
    if ($Params.ContainsKey('HardReset')) { $hardReset = [bool]$Params.HardReset }
    if ($Params.ContainsKey('BrowserDataReset')) { $browserDataReset = [bool]$Params.BrowserDataReset }
    if ($hardReset) { $browserDataReset = $false }

    $result = Invoke-ChromeReset -BrowserDataReset:$browserDataReset -HardReset:$hardReset
    exit $result.ExitCode
}

# ============================================================
# GUI (nur Interactive)
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form               = New-Object System.Windows.Forms.Form
$Form.Text          = "Google Chrome - Reset"
$Form.StartPosition = "CenterScreen"
$Form.Size          = New-Object System.Drawing.Size(780, 460)
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
    $style = if ($Bold) {[System.Drawing.FontStyle]::Bold} else {[System.Drawing.FontStyle]::Regular}
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $style)
    $label.Text = $Text
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(730, 0)
    $label.Location = New-Object System.Drawing.Point(10, $script:y)
    $label.ForeColor = $Color
    $Panel.Controls.Add($label)
    $script:y += $label.Height + 8
}

Add-Label "Google Chrome - Reset" 12 $true
Add-Label "Browserdaten loeschen: entfernt Cache, Cookies und Website-Daten fuer alle Chrome-Profile. Profile und Lesezeichen bleiben erhalten (empfohlen als erster Schritt)." 10 $false ([System.Drawing.Color]::DarkRed)
Add-Label "Hard-Reset: setzt Chrome komplett zurueck. Der komplette Chrome-Ordner in AppData wird geloescht. Lesezeichen und 'First Run' werden gesichert und wiederhergestellt." 10 $false ([System.Drawing.Color]::DarkRed)

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $script:y)
$grp.Size = New-Object System.Drawing.Size(740, 130)
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

$cbBrowser = Add-Check "Browserdaten loeschen (Cache + Cookies + Website-Daten) - empfohlen"
$cbHard    = Add-Check "Hard-Reset: Chrome komplett zuruecksetzen (AppData + Registry, Bookmarks bleiben)"

$cbBrowser.Checked = $true
$cbHard.Checked    = $false

$cbHard.Add_CheckedChanged({
    if ($cbHard.Checked) {
        $cbBrowser.Checked = $false
        $cbBrowser.Enabled = $false
    } else {
        $cbBrowser.Enabled = $true
        $cbBrowser.Checked = $true
    }
})

$script:y += 150

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Reset ausfuehren"
$btnRun.Width = 140
$btnRun.Location = New-Object System.Drawing.Point(20, $script:y)
$Panel.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Schliessen"
$btnClose.Width = 100
$btnClose.Location = New-Object System.Drawing.Point(180, $script:y)
$Panel.Controls.Add($btnClose)
$btnClose.Add_Click({ $Form.Close() })

$Form.AcceptButton = $btnRun
$Form.CancelButton = $btnClose

$btnRun.Add_Click({
    if (-not $cbBrowser.Checked -and -not $cbHard.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte mindestens eine Aktion auswaehlen.",
            "Hinweis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ($cbHard.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG: Hard-Reset entfernt alle Chrome-Daten in AppData und setzt Chrome vollstaendig zurueck.`nLesezeichen und 'First Run' werden gesichert und wiederhergestellt.`n`nFortfahren?",
            "Hard-Reset bestaetigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    } elseif ($cbBrowser.Checked) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Browserdaten (Cache, Cookies und Website-Daten) werden geloescht. Sie werden auf Webseiten abgemeldet.`n`nFortfahren?",
            "Browserdaten loeschen bestaetigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $btnRun.Enabled             = $false
    $script:overallSuccess      = $true
    $script:bookmarkBackupOk    = $false

    try {
        $result = Invoke-ChromeReset `
            -BrowserDataReset:$cbBrowser.Checked `
            -HardReset:$cbHard.Checked

        [System.Windows.Forms.MessageBox]::Show("Reset abgeschlossen.", "Fertig", 'OK', 'Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler beim Ausfuehren des Resets:`n$($_.Exception.Message)", "Fehler", 'OK', 'Error'
        ) | Out-Null
    } finally {
        $btnRun.Enabled = $true
    }
})

[void]$Form.ShowDialog()
Write-Log "=== Google Chrome Reset beendet ==="
