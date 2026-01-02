Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================
# Logging (Fallback, falls nicht vom Hauptskript bereitgestellt)
# ==========================
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
        )
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "[$ts] [$Level] $Message"
    }
}

Write-Log "=== Outlook Reset (GUI) gestartet ==="

# ==========================
# Sichere Pfad-Löschung
# ==========================
function Remove-PathSafe {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Nicht vorhanden: {0}" -f $Path)
        return $true
    }

    # Native Commands sauber ausführen, auch wenn Script von UNC gestartet wurde
    $origLoc = Get-Location
    try { Set-Location -LiteralPath $env:SystemRoot } catch {}

    try {
        # 1) Erst per cmd rd (schnell/robust) – Output und Errors IN CMD unterdrücken
        try { & cmd.exe /d /c "rd /s /q ""$Path"" >nul 2>nul" | Out-Null } catch {}

        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log ("Gelöscht: {0}" -f $Path)
            return $true
        }

        # 2) Robocopy-Mirror mit leerem Ordner (leert WebView2/CacheStorage sehr zuverlässig)
        $empty = Join-Path $env:TEMP ("_empty_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $empty -Force | Out-Null
        try {
            & robocopy.exe $empty $Path /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP > $null 2> $null
        } catch {
            Write-Log ("Robocopy-MIR fehlgeschlagen: {0}" -f $_.Exception.Message) 'WARN'
        }
        try { Remove-Item -LiteralPath $empty -Recurse -Force -ErrorAction SilentlyContinue } catch {}

        # 3) Danach erneut rd + Remove-Item (alles still)
        try { & cmd.exe /d /c "rd /s /q ""$Path"" >nul 2>nul" | Out-Null } catch {}
        if (Test-Path -LiteralPath $Path) {
            try { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log ("Gelöscht: {0}" -f $Path)
            return $true
        }

        Write-Log ("Pfad konnte nicht vollständig entfernt werden (Rest vorhanden): {0}" -f $Path) 'WARN'
        return $false
    }
    catch {
        # "Ein Teil des Pfades..." ist bei WebView2/CacheStorage häufig ein Race-Condition-Effekt
        $msg = $_.Exception.Message
        if ($msg -match 'konnte nicht gefunden werden|could not be found|Ein Teil des Pfades') {
            if (-not (Test-Path -LiteralPath $Path)) {
                Write-Log ("Gelöscht (trotz PathNotFound während der Bereinigung): {0}" -f $Path) 'WARN'
                return $true
            }
            Write-Log ("PathNotFound während Bereinigung, Rest evtl. vorhanden: {0}" -f $Path) 'WARN'
            return $false
        }

        if ($msg -like '*von einem anderen Prozess verwendet*') {
            Write-Log ("Pfad in Benutzung, übersprungen: {0}" -f $Path) 'WARN'
            return $false
        }

        Write-Log ("Fehler beim Löschen: {0} -> {1}" -f $Path, $msg) 'ERROR'
        return $false
    }
    finally {
        try { Set-Location -LiteralPath $origLoc.Path } catch {}
    }
}

# ==========================
# Office-/Search-Prozesse beenden (nur User-Session)
# ==========================
function Stop-OfficeProcesses {
    [CmdletBinding()]
    param(
        [string[]]$ProcessNames = @(
            # Outlook/Office
            'outlook','winword','excel','powerpnt','onenote','msaccess','publisher','visio','project',
            # Office Shared / Add-in related (Userkontext)
            'msoia','olk',
            # Search Hosts (User-Session)
            'searchprotocolhost','searchfilterhost'
        ),
        [int]$Retries = 8,
        [int]$DelayMs = 350
    )

    try {
        $sessionId = (Get-Process -Id $PID).SessionId
    }
    catch {
        Write-Log ("Konnte SessionId nicht ermitteln → beende ohne Session-Filter: {0}" -f $_.Exception.Message) 'WARN'
        $sessionId = $null
    }

    $lower = $ProcessNames | ForEach-Object { $_.ToLower() }
    Write-Log ("Beende Office-/Outlook-Prozesse{0} …" -f $(if ($sessionId -ne $null) { " in Session $sessionId" } else { "" }))

    for ($i = 1; $i -le $Retries; $i++) {
        try {
            $candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                ($lower -contains $_.Name.ToLower()) -and
                (($sessionId -eq $null) -or ($_.SessionId -eq $sessionId))
            }

            foreach ($p in $candidates) {
                try {
                    Stop-Process -Id $p.Id -Force -ErrorAction Stop
                    Write-Log ("Prozess beendet: {0} (PID {1})" -f $p.Name, $p.Id)
                } catch {
                    Write-Log ("Fehler beim Beenden von {0} (PID {1}): {2}" -f $p.Name,$p.Id,$_.Exception.Message) 'WARN'
                }
            }
        }
        catch {
            Write-Log ("Fehler beim Prozess-Stopp: {0}" -f $_.Exception.Message) 'WARN'
        }

        Start-Sleep -Milliseconds $DelayMs
    }

    # Kurz prüfen ob noch etwas läuft
    $left = Get-Process -ErrorAction SilentlyContinue | Where-Object {
        ($lower -contains $_.Name.ToLower()) -and
        (($sessionId -eq $null) -or ($_.SessionId -eq $sessionId))
    }

    if ($left) {
        Write-Log ("Laufen noch: {0}" -f (($left | Select-Object -ExpandProperty Name -Unique) -join ', ')) 'WARN'
        return $false
    }

    Write-Log "Office-/Outlook-Prozesse beendet."
    return $true
}

# ==========================
# Cache löschen (Standard-Reset)
# ==========================
function Clear-OutlookCache {
    Write-Log "Starte Standard-Reset (Cache & temporäre Dateien)…"

    $paths = @(
        (Join-Path $env:APPDATA      'Microsoft\Outlook\RoamCache'),
        (Join-Path $env:APPDATA      'Microsoft\Outlook\Offline Address Books'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCache\Content.Outlook')
    )

    foreach ($p in $paths) {
        try {
            if (Test-Path $p) {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Log "Gelöscht: $p"
            } else {
                Write-Log "Nicht vorhanden: $p"
            }
        }
        catch {
            Write-Log ("Konnte Cachepfad nicht löschen: {0} → {1}" -f $p, $_.Exception.Message) 'WARN'
        }
    }

    Write-Log "Standard-Reset abgeschlossen."
}

# ==========================
# Profil-Reset (inkl. Cache)
# ==========================
function Reset-OutlookProfiles {
    Write-Log "Starte Profil-Reset …"

    # Profil-Reset beinhaltet immer Cache-Reset
    Clear-OutlookCache

    $regKeys = @(
        'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles',
        'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover',
        'HKCU:\Software\Microsoft\Office\16.0\Common\Identity'
    )

    foreach ($key in $regKeys) {
        try {
            if (Test-Path $key) {
                Remove-Item -Path $key -Recurse -Force -ErrorAction Stop
                Write-Log "Registry gelöscht: $key"
            } else {
                Write-Log "Registry nicht vorhanden: $key"
            }
        }
        catch {
            Write-Log ("Fehler beim Löschen der Registry {0}: {1}" -f $key, $_.Exception.Message) 'WARN'
        }
    }

    Write-Log "Profil-Reset abgeschlossen."
}

# ==========================
# Hard-Reset: kompletter Outlook-Schlüssel
# ==========================
function Reset-OutlookHard {
    Write-Log "Starte Hard-Reset: kompletter Outlook-Registry-Schlüssel…"

    # Hard-Reset beinhaltet ebenfalls Cache-Reset
    Clear-OutlookCache

    $root = 'HKCU:\Software\Microsoft\Office\16.0\Outlook'

    try {
        if (Test-Path $root) {
            Remove-Item -Path $root -Recurse -Force -ErrorAction Stop
            Write-Log "Kompletter Outlook-Schlüssel gelöscht: $root"
        } else {
            Write-Log "Outlook-Schlüssel nicht vorhanden: $root"
        }
    }
    catch {
        Write-Log ("Fehler beim Löschen des Outlook-Schlüssels: {0}" -f $_.Exception.Message) 'WARN'
    }

    Write-Log "Hard-Reset abgeschlossen."
}

# ==========================
# Office Web Add-ins reparieren (WebView2/Wef)
# ==========================
function Repair-OfficeWebAddins {
    Write-Log "Starte Reparatur: Office-Web-Add-Ins (Wef/WebView2)…"

    $wefWebView2 = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef\webview2'
    $deleted = Remove-PathSafe -Path $wefWebView2

    if ($deleted) {
        Write-Log ("Office Web Add-ins: Ordner bereinigt: {0}" -f $wefWebView2)
    } else {
        Write-Log ("Office Web Add-ins: Ordner konnte nicht vollständig bereinigt werden: {0}" -f $wefWebView2) 'WARN'
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

    Write-Log "Outlook-Reset gestartet …"

    # Nur wenn wirklich eine Aktion gewählt ist, die Office schließen muss
    if ($StandardReset -or $ProfileReset -or $HardReset -or $RepairWebAddins) {
        Stop-OfficeProcesses | Out-Null
    }

    if ($HardReset) {
        Reset-OutlookHard
    }
    elseif ($ProfileReset) {
        Reset-OutlookProfiles
    }
    elseif ($StandardReset) {
        Clear-OutlookCache
    }

    if ($RepairWebAddins) {
        Repair-OfficeWebAddins
    }

    Write-Log "Outlook-Reset abgeschlossen."
    return @{ ExitCode = 0; Errors = 0; Warnings = 0 }
}

# ==========================
# GUI
# ==========================
$Form               = New-Object System.Windows.Forms.Form
$Form.Text          = "Microsoft Outlook – Reset"
$Form.StartPosition = "CenterScreen"
$Form.Size          = New-Object System.Drawing.Size(750, 500)
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
    $label.MaximumSize = New-Object System.Drawing.Size(700, 0)
    $label.Location = New-Object System.Drawing.Point(10, $script:y)
    $label.ForeColor = $Color
    $Panel.Controls.Add($label)
    $script:y += $label.Height + 8
}

Add-Label "Microsoft Outlook – Reset" 12 $true
Add-Label "Standard-Reset: löscht den Outlook-Cache/ temporäre Dateien (empfohlen)." 10 $false
Add-Label "Profil-Reset: setzt das Outlook-Profil zurück (inkl. Cache). Outlook-Profil muss danach neu erstellt werden." 10 $false
Add-Label "Office-Web-Add-Ins reparieren: löscht %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2 (Office-Apps werden dafür beendet)." 10 $false
Add-Label "Hard-Reset: Alle Outlook Einstellungen zurücksetzten (inkl. Cache/ Profil) verhält sich wie bei einem neuen Benutzerprofil" 10 $false ([System.Drawing.Color]::DarkRed)

# GroupBox für Aktionen
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $y)
$grp.Size = New-Object System.Drawing.Size(710, 200)
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

$cbStandard   = Add-Check "Standard-Reset: Cache löschen (empfohlen)"
$cbProfile    = Add-Check "Profil-Reset: Outlook-Profil zurücksetzen (inkl. Cache)"
$cbWebAddins  = Add-Check "Office-Web-Add-Ins reparieren (Wef/WebView2 Cache löschen)"
$cbHard       = Add-Check "Hard-Reset: Alle Outlook Einstellungen zurücksetzten (inkl. Cache/ Profil)"

# Standard: nur Standard-Reset aktiv
$cbStandard.Checked  = $false
$cbProfile.Checked   = $false
$cbHard.Checked      = $false
$cbWebAddins.Checked = $false

# Abhängigkeit Profil → Standard (sofern kein Hard-Reset aktiv ist)
$cbProfile.Add_CheckedChanged({
    if ($cbHard.Checked) { return }  # Hard-Reset steuert alles selbst
    if ($cbProfile.Checked) {
        $cbStandard.Checked = $true
        $cbStandard.Enabled = $false
    } else {
        $cbStandard.Enabled = $true
        $cbStandard.Checked = $true
    }
})

# Hard-Reset → erzwingt Profil + Standard, deaktiviert beide
$cbHard.Add_CheckedChanged({
    if ($cbHard.Checked) {
        $cbProfile.Checked = $true
        $cbProfile.Enabled = $false

        $cbStandard.Checked = $true
        $cbStandard.Enabled = $false
    } else {
        $cbProfile.Enabled = $true
        $cbStandard.Enabled = $true
        $cbStandard.Checked = $true
    }
})

$y += 220

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
    if (-not $cbStandard.Checked -and -not $cbProfile.Checked -and -not $cbHard.Checked -and -not $cbWebAddins.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte mindestens eine Aktion auswählen.",
            "Hinweis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    # Starke Warnung bei Hard-Reset
    if ($cbHard.Checked) {
        $resHard = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG: Beim Hard-Reset wird der komplette Outlook-Registry-Bereich dieses Benutzers gelöscht." +
            "`nAlle Outlook-Einstellungen, Profile, Ansichten etc. werden zurückgesetzt." +
            "`nOutlook verhält sich danach wie frisch installiert." +
            "`n`nFortfahren?",
            "Hard-Reset bestätigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($resHard -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    elseif ($cbProfile.Checked) {
        $resProf = [System.Windows.Forms.MessageBox]::Show(
            "Outlook-Profil wird gelöscht und muss neu eingerichtet werden." +
            "`nFortfahren?",
            "Profil-Reset bestätigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($resProf -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    # Hinweis bei WebAddins-Reparatur: Office Apps werden beendet
    if ($cbWebAddins.Checked) {
        $resWef = [System.Windows.Forms.MessageBox]::Show(
            "Für die Reparatur der Office-Web-Add-Ins müssen alle Office-Anwendungen beendet werden (Outlook, Word, Excel, PowerPoint etc.)." +
            "`nOffene Dokumente bitte vorher speichern." +
            "`n`nFortfahren?",
            "Office-Web-Add-Ins reparieren",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($resWef -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $btnRun.Enabled = $false
    try {
        Invoke-OutlookReset `
            -StandardReset:$cbStandard.Checked `
            -ProfileReset:$cbProfile.Checked `
            -HardReset:$cbHard.Checked `
            -RepairWebAddins:$cbWebAddins.Checked

        [System.Windows.Forms.MessageBox]::Show(
            "Aktion(en) abgeschlossen.",
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
Write-Log "=== Outlook Reset (GUI) beendet ==="