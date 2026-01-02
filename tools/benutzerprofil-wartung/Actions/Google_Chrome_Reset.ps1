[CmdletBinding()]
param(
    [ValidateSet('Interactive','Silent')]
    [string]$Mode = 'Interactive',
    [hashtable]$Params = @{}
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot 'shared\WartungsTools.SDK.psm1') -Force
$toolId = (Get-Content (Join-Path $toolRoot 'tool.json') -Raw | ConvertFrom-Json).toolId
$actionName = 'Google_Chrome_Reset'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    WartungsTools.SDK\Write-Log -Level $Level -Message $Message -ToolId $toolId -Action $actionName
}

if ($Mode -eq 'Interactive') {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
# 0) Logging-Fallback (falls Write-Log vom Hauptskript nicht existiert)
# ============================================================
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO','WARN','ERROR')]
            [string]$Level = 'INFO'
        )

        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $prefix = "[ChromeReset][$Level]"

        switch ($Level) {
            'INFO'  { Write-Host "$prefix $Message" -ForegroundColor Gray }
            'WARN'  { Write-Host "$prefix $Message" -ForegroundColor Yellow }
            'ERROR' { Write-Host "$prefix $Message" -ForegroundColor Red }
        }
    }
}

Write-Log "=== Google Chrome Reset (GUI) gestartet ==="

# globale Statusvariablen
$script:overallSuccess   = $true
$script:bookmarkBackupOk = $false

# Standardpfade
$script:ChromeUserDataRoot   = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
$script:ChromeDefaultPath    = Join-Path $script:ChromeUserDataRoot 'Default'
# Backup außerhalb des Chrome-Ordners, damit Hard-Reset ihn nicht löscht
$script:BackupFolder         = Join-Path $env:LOCALAPPDATA 'Wartungsscript\ChromeRestore'

# ============================================================
# 1) Chrome-Prozesse nur im aktuellen User-Session-Kontext beenden
# ============================================================
function Stop-ChromeProcesses {
    Write-Log "Beende Chrome-Prozesse des aktuellen Benutzers..."

    try {
        $sessionId = (Get-Process -Id $PID).SessionId
    }
    catch {
        Write-Log ("Konnte SessionId nicht ermitteln → beende ohne Filter: {0}" -f $_.Exception.Message) 'WARN'
        $sessionId = $null
    }

    $targets = @("chrome", "googleupdate", "googlecrashhandler")

    for ($i = 1; $i -le 8; $i++) {
        try {
            if ($sessionId -ne $null) {
                Get-Process -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -in $targets -and $_.SessionId -eq $sessionId } |
                    Stop-Process -Force -ErrorAction SilentlyContinue
            }
            else {
                Get-Process -Name $targets -ErrorAction SilentlyContinue |
                    Stop-Process -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log ("Fehler beim Prozessstopp: {0}" -f $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }

        Start-Sleep -Milliseconds 300
    }

    Write-Log "Chrome-Prozesse beendet."
}

# ============================================================
# 2) Browserdaten löschen (Cache + Cookies + Website-Daten, alle Profile)
# ============================================================
function Clear-ChromeBrowserData {
    Write-Log "Starte Löschung von Browserdaten (Cache + Cookies + Website-Daten) für alle Chrome-Profile …"

    if (-not (Test-Path $script:ChromeUserDataRoot)) {
        Write-Log ("Chrome User Data Pfad nicht vorhanden: {0}" -f $script:ChromeUserDataRoot) 'WARN'
        return
    }

    $profileDirs = Get-ChildItem -Path $script:ChromeUserDataRoot -Directory -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' -or $_.Name -in @('Guest Profile','System Profile') }

    # Cache-Ordner
    $cacheSubDirs = @(
        'Cache',
        'Code Cache',
        'GPUCache',
        'Media Cache',
        'ShaderCache',
        'Service Worker\CacheStorage',
        'Service Worker\ScriptCache',
        'Storage\ext'
    )

    # Cookie-Dateien
    $cookieFiles = @(
        'Cookies',
        'Cookies-journal',
        'Network\Cookies',
        'Network\Cookies-journal'
    )

    # Website-/Site-Daten
    $siteDirs = @(
        'Local Storage',
        'Local Storage\leveldb',
        'Session Storage',
        'IndexedDB',
        'File System',
        'Service Worker',
        'Service Worker\CacheStorage',
        'Service Worker\Database',
        'Service Worker\ScriptCache',
        'Storage'
    )

    foreach ($dir in $profileDirs) {

        # Cache
        foreach ($sub in $cacheSubDirs) {
            $p = Join-Path $dir.FullName $sub
            try {
                if (Test-Path $p) {
                    Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                    Write-Log ("Cache-Ordner gelöscht: {0}" -f $p)
                }
            }
            catch {
                Write-Log ("Fehler beim Löschen von Cache-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }

        # Cookies
        foreach ($cf in $cookieFiles) {
            $p = Join-Path $dir.FullName $cf
            try {
                if (Test-Path $p) {
                    Remove-Item -Path $p -Force -ErrorAction Stop
                    Write-Log ("Cookies-Datei gelöscht: {0}" -f $p)
                }
            }
            catch {
                Write-Log ("Fehler beim Löschen von Cookies-Datei {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }

        # Site-Daten
        foreach ($sd in $siteDirs) {
            $p = Join-Path $dir.FullName $sd
            try {
                if (Test-Path $p) {
                    Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                    Write-Log ("Site-Daten-Ordner gelöscht: {0}" -f $p)
                }
            }
            catch {
                Write-Log ("Fehler beim Löschen von Site-Daten-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }
    }

    # einige globale Cache-Strukturen direkt unter User Data
    $globalCacheDirs = @('Crashpad', 'SwReporter', 'ShaderCache')
    foreach ($g in $globalCacheDirs) {
        $p = Join-Path $script:ChromeUserDataRoot $g
        try {
            if (Test-Path $p) {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Log ("Globaler Cache-Ordner gelöscht: {0}" -f $p)
            }
        }
        catch {
            Write-Log ("Fehler beim Löschen von globalem Cache-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }

    Write-Log "Löschung von Browserdaten (Cache + Cookies + Website-Daten) abgeschlossen."
}

# ============================================================
# 3) Chrome-Registry (HKCU) bereinigen
# ============================================================
function Clear-ChromeRegistry {
    Write-Log "Bereinige Chrome-Registry-Einträge (HKCU)..."

    $chromeRootKey = "HKCU:\Software\Google\Chrome"

    # Wenn der Schlüssel gar nicht existiert → fertig
    if (-not (Test-Path $chromeRootKey)) {
        Write-Log ("Registry nicht vorhanden: {0}" -f $chromeRootKey)
        return
    }

    # --- Unterschlüssel rekursiv löschen ---
    try {
        Get-ChildItem -Path $chromeRootKey -ErrorAction SilentlyContinue |
            Sort-Object -Property PSPath -Descending |
            ForEach-Object {
                try {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Log ("Registry-Unterzweig gelöscht: {0}" -f $_.Name)
                }
                catch {
                    Write-Log ("Fehler beim Löschen von Unterzweig {0}: {1}" -f $_.Name, $_.Exception.Message) 'WARN'
                    $script:overallSuccess = $false
                }
            }
    }
    catch {
        Write-Log ("Fehler beim Auflisten der Chrome-Unterschlüssel: {0}" -f $_.Exception.Message) 'WARN'
        $script:overallSuccess = $false
    }

    # --- Werte direkt unter HKCU\Software\Google\Chrome löschen ---
    try {
        $props = Get-ItemProperty -Path $chromeRootKey -ErrorAction Stop

        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -in 'PSPath','PSParentPath','PSChildName','PSDrive','PSProvider') { continue }

            try {
                Remove-ItemProperty -Path $chromeRootKey -Name $prop.Name -ErrorAction SilentlyContinue
                Write-Log ("Registry-Wert gelöscht: {0} -> {1}" -f $chromeRootKey, $prop.Name)
            }
            catch {
                Write-Log ("Fehler beim Löschen von Wert {0} in {1}: {2}" -f $prop.Name, $chromeRootKey, $_.Exception.Message) 'WARN'
                $script:overallSuccess = $false
            }
        }
    }
    catch {
        Write-Log ("Fehler beim Lesen der Werte unter {0}: {1}" -f $chromeRootKey, $_.Exception.Message) 'WARN'
        $script:overallSuccess = $false
    }

    # --- Root-Schlüssel optional entfernen ---
    try {
        Remove-Item -Path $chromeRootKey -Force -Recurse -ErrorAction SilentlyContinue -Confirm:$false

        if (-not (Test-Path $chromeRootKey)) {
            Write-Log ("Chrome-Root-Schlüssel entfernt: {0}" -f $chromeRootKey)
        }
        else {
            Write-Log ("Chrome-Root-Schlüssel konnte nicht entfernt werden (Policy/Berechtigung), Unterstruktur bereinigt: {0}" -f $chromeRootKey) 'INFO'
        }
    }
    catch {
        Write-Log ("Chrome-Root-Schlüssel konnte nicht entfernt werden (Policy/Berechtigung). Unterstruktur wurde bereinigt: {0}" -f $chromeRootKey) 'INFO'
    }
}

# ============================================================
# 4) Bookmarks + "First Run" sichern / wiederherstellen
# ============================================================
function Backup-ChromeBookmarks {
    $script:bookmarkBackupOk = $false

    $bookmarksFile = Join-Path $script:ChromeDefaultPath 'Bookmarks'
    $firstRunFile  = Join-Path $script:ChromeUserDataRoot 'First Run'

    if (-not (Test-Path $bookmarksFile)) {
        Write-Log ("Keine Bookmarks-Datei gefunden: {0}" -f $bookmarksFile) 'WARN'
    }

    if (-not (Test-Path $bookmarksFile) -and -not (Test-Path $firstRunFile)) {
        # nichts zu sichern
        return
    }

    if (-not (Test-Path $script:BackupFolder)) {
        try {
            New-Item -ItemType Directory -Path $script:BackupFolder -ErrorAction Stop | Out-Null
            Write-Log ("Backup-Ordner erstellt: {0}" -f $script:BackupFolder)
        }
        catch {
            Write-Log ("Fehler beim Erstellen des Backup-Ordners: {0}" -f $_.Exception.Message) 'ERROR'
            $script:overallSuccess = $false
            return
        }
    }

    # Bookmarks sichern (wenn vorhanden)
    if (Test-Path $bookmarksFile) {
        try {
            Copy-Item -Path $bookmarksFile -Destination $script:BackupFolder -Force
            Write-Log ("Lesezeichen wurden nach {0} gesichert." -f $script:BackupFolder)
            $script:bookmarkBackupOk = $true
        }
        catch {
            Write-Log ("Fehler beim Sichern der Lesezeichen: {0}" -f $_.Exception.Message) 'ERROR'
            $script:overallSuccess   = $false
            $script:bookmarkBackupOk = $false
        }
    }

    # "First Run" sichern (wenn vorhanden)
    if (Test-Path $firstRunFile) {
        try {
            Copy-Item -Path $firstRunFile -Destination (Join-Path $script:BackupFolder 'First Run') -Force
            Write-Log ("'First Run'-Datei wurde gesichert: {0}" -f $firstRunFile)
        }
        catch {
            Write-Log ("Fehler beim Sichern der 'First Run'-Datei: {0}" -f $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }
}

function Restore-ChromeBookmarks {
    # User Data Root sicherstellen
    if (-not (Test-Path $script:ChromeUserDataRoot)) {
        try {
            New-Item -ItemType Directory -Path $script:ChromeUserDataRoot -ErrorAction Stop | Out-Null
            Write-Log ("Chrome User Data Root neu erstellt: {0}" -f $script:ChromeUserDataRoot)
        }
        catch {
            Write-Log ("Fehler beim Erstellen von Chrome User Data Root: {0}" -f $_.Exception.Message) 'ERROR'
            $script:overallSuccess = $false
            return
        }
    }

    # "First Run" wiederherstellen (unabhängig von Bookmarks)
    $backupFirstRun = Join-Path $script:BackupFolder 'First Run'
    if (Test-Path $backupFirstRun) {
        try {
            Copy-Item -Path $backupFirstRun -Destination (Join-Path $script:ChromeUserDataRoot 'First Run') -Force
            Write-Log ("'First Run'-Datei wurde wiederhergestellt: {0}" -f (Join-Path $script:ChromeUserDataRoot 'First Run'))
        }
        catch {
            Write-Log ("Fehler beim Wiederherstellen der 'First Run'-Datei: {0}" -f $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }

    # Bookmarks nur, wenn Backup erfolgreich war
    if (-not $script:bookmarkBackupOk) {
        Write-Log "Bookmark-Backup war nicht erfolgreich – keine Wiederherstellung der Lesezeichen möglich." 'WARN'
        return
    }

    $backupBookmarksFile = Join-Path $script:BackupFolder 'Bookmarks'

    if (-not (Test-Path $backupBookmarksFile)) {
        Write-Log ("Gesicherte Bookmarks-Datei nicht gefunden: {0}" -f $backupBookmarksFile) 'WARN'
        return
    }

    try {
        if (-not (Test-Path $script:ChromeDefaultPath)) {
            New-Item -ItemType Directory -Path $script:ChromeDefaultPath -ErrorAction Stop | Out-Null
            Write-Log ("Neuen Default-Ordner erstellt: {0}" -f $script:ChromeDefaultPath)
        }

        Copy-Item -Path $backupBookmarksFile -Destination $script:ChromeDefaultPath -Force
        Write-Log ("Lesezeichen wurden erfolgreich wiederhergestellt nach: {0}" -f $script:ChromeDefaultPath)
    }
    catch {
        Write-Log ("Fehler beim Wiederherstellen der Lesezeichen: {0}" -f $_.Exception.Message) 'ERROR'
        $script:overallSuccess = $false
    }
}

function Cleanup-BookmarkBackup {
    if (Test-Path $script:BackupFolder) {
        try {
            Remove-Item -Path $script:BackupFolder -Recurse -Force -ErrorAction Stop
            Write-Log ("Bookmark-Backup-Ordner gelöscht: {0}" -f $script:BackupFolder)
        }
        catch {
            Write-Log ("Fehler beim Löschen des Bookmark-Backups: {0}" -f $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }
}

# ============================================================
# 5) Hard-Reset (kompletter Chrome-Baum in AppData)
# ============================================================
function Reset-ChromeHard {
    Write-Log "Starte Hard-Reset für Chrome…"

    Backup-ChromeBookmarks

    # kompletter Chrome-Ordner in Local + Roaming AppData
    $localChrome   = Join-Path $env:LOCALAPPDATA 'Google\Chrome'
    $roamingChrome = Join-Path $env:APPDATA      'Google\Chrome'

    foreach ($p in @($localChrome, $roamingChrome)) {
        try {
            if (Test-Path $p) {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Log ("Chrome-Ordner gelöscht: {0}" -f $p)
            }
            else {
                Write-Log ("Chrome-Ordner nicht vorhanden: {0}" -f $p)
            }
        }
        catch {
            Write-Log ("Fehler beim Löschen von Chrome-Ordner {0}: {1}" -f $p, $_.Exception.Message) 'WARN'
            $script:overallSuccess = $false
        }
    }

    Clear-ChromeRegistry

    # Pfade neu setzen, falls User Data Root gelöscht wurde
    $script:ChromeUserDataRoot = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
    $script:ChromeDefaultPath  = Join-Path $script:ChromeUserDataRoot 'Default'

    Restore-ChromeBookmarks
    Cleanup-BookmarkBackup

    Write-Log "Hard-Reset abgeschlossen."
}

# ============================================================
# 6) Hauptfunktion: Invoke-ChromeReset
# ============================================================
function Invoke-ChromeReset {
    [CmdletBinding()]
    param(
        [switch]$BrowserDataReset,
        [switch]$HardReset
    )

    Write-Log "Chrome-Reset gestartet …"

    Stop-ChromeProcesses

    if ($HardReset) {
        Reset-ChromeHard
    }
    elseif ($BrowserDataReset) {
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
# 7) GUI
# ============================================================
$Form               = New-Object System.Windows.Forms.Form
$Form.Text          = "Google Chrome – Reset"
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

Add-Label "Google Chrome – Reset" 12 $true
Add-Label "Browserdaten löschen: entfernt Cache, Cookies und Website-Daten für alle Chrome-Profile. Profile und Lesezeichen bleiben erhalten (empfohlen als erster Schritt)." 10 $false ([System.Drawing.Color]::DarkRed)
Add-Label "Hard-Reset: setzt Chrome komplett zurück. Der komplette Chrome-Ordner in AppData (Local + Roaming) wird gelöscht und neu erstellt. Lesezeichen aus dem Standardprofil sowie die 'First Run'-Information werden gesichert und wiederhergestellt." 10 $false ([System.Drawing.Color]::DarkRed)

# GroupBox für Aktionen
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $y)
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

$cbBrowser = Add-Check "Browserdaten löschen (Cache + Cookies + Website-Daten) – empfohlen"
$cbHard    = Add-Check "Hard-Reset: Chrome komplett zurücksetzen (AppData + Registry, Bookmarks + 'First Run' bleiben)"

# Standard: nur Browserdaten aktiv
$cbBrowser.Checked = $true
$cbHard.Checked    = $false

# Abhängigkeit: Hard-Reset schließt Browserdaten-Option aus (ist ohnehin inkludiert)
$cbHard.Add_CheckedChanged({
    if ($cbHard.Checked) {
        $cbBrowser.Checked = $false
        $cbBrowser.Enabled = $false
    } else {
        $cbBrowser.Enabled = $true
        $cbBrowser.Checked = $true
    }
})

$y += 150

# Buttons
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Reset ausführen"
$btnRun.Width = 140
$btnRun.Location = New-Object System.Drawing.Point(20, $y)
$Panel.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Schließen"
$btnClose.Width = 100
$btnClose.Location = New-Object System.Drawing.Point(180, $y)
$Panel.Controls.Add($btnClose)
$btnClose.Add_Click({ $Form.Close() })

$Form.AcceptButton = $btnRun
$Form.CancelButton = $btnClose

$btnRun.Add_Click({
    if (-not $cbBrowser.Checked -and -not $cbHard.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte mindestens eine Aktion auswählen.",
            "Hinweis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    # Bestätigungen
    if ($cbHard.Checked) {
        $resHard = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG: Hard-Reset entfernt alle Chrome-Daten in AppData (Local + Roaming) und setzt Chrome vollständig zurück." +
            "`nLesezeichen aus dem Standardprofil sowie die 'First Run'-Information werden gesichert und wiederhergestellt." +
            "`n`nFortfahren?",
            "Hard-Reset bestätigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($resHard -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }
    elseif ($cbBrowser.Checked) {
        $resBrowser = [System.Windows.Forms.MessageBox]::Show(
            "Browserdaten (Cache, Cookies und Website-Daten) werden gelöscht. Dies führt in der Regel dazu, dass Sie auf Webseiten abgemeldet werden." +
            "`n`nFortfahren?",
            "Browserdaten löschen bestätigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($resBrowser -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    $btnRun.Enabled             = $false
    $script:overallSuccess      = $true
    $script:bookmarkBackupOk    = $false

    try {
        $result = Invoke-ChromeReset `
            -BrowserDataReset:$cbBrowser.Checked `
            -HardReset:$cbHard.Checked

        [System.Windows.Forms.MessageBox]::Show(
            "Reset abgeschlossen.",
            "Fertig",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler beim Ausführen des Resets:`n$($_.Exception.Message)",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $btnRun.Enabled = $true
    }
})

[void]$Form.ShowDialog()
Write-Log "=== Google Chrome Reset (GUI) beendet ==="
}
}

