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

if ($Mode -eq 'Interactive') {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
# ==========================
# Prozesse beenden
# ==========================
function Stop-UserProcesses {
    [CmdletBinding()]
    param(
        [string[]]$ProcessNames = @(
            # Teams / WebView2
            'ms-teams','teams','teamswebview2','msedgewebview2',
            'teams-updater','teamsbootstrapper','update','teamsupdate',

            # Auth / AAD Broker / Token Handling (Userkontext)
            'microsoft.aad.brokerplugin',
            'oneauth',
            'tokenbroker'
        )
    )

    $sessionId = (Get-Process -Id $PID).SessionId
    $killed = 0
    Write-Log ("Suche Zielprozesse in Session {0}: {1}" -f $sessionId, ($ProcessNames -join ', '))

    try {
        Write-Log "Prozessliste vor Kill:"
        Get-Process | Where-Object { $_.Name -match 'teams|webview2|oneauth|tokenbroker|aad' } |
            Select-Object Name, Id, SessionId |
            ForEach-Object { Write-Log ("  {0} (PID {1}, Session {2})" -f $_.Name,$_.Id,$_.SessionId) }

        # 1. Kandidaten in der aktuellen Session suchen
        $candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
            ($ProcessNames -contains $_.Name.ToLower()) -and ($_.SessionId -eq $sessionId)
        }

        if ($candidates) {
            Write-Log ("Gefunden: {0}" -f (($candidates | ForEach-Object { "$($_.Name)/$($_.Id)" }) -join ', '))

            # 2. Normaler Stop-Process
            foreach ($p in $candidates) {
                try {
                    Stop-Process -Id $p.Id -Force -ErrorAction Stop
                    Write-Log ("Stop-Process OK: {0} (PID {1})" -f $p.Name, $p.Id)
                    $killed++
                } catch {
                    Write-Log ("Stop-Process fehlgeschlagen für {0} (PID {1}): {2}" -f $p.Name,$p.Id,$_.Exception.Message) 'WARN'
                }
            }

            # 3. Was noch lebt, mit taskkill /F /T erwischen
            $stillRunning = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                ($ProcessNames -contains $_.Name.ToLower()) -and ($_.SessionId -eq $sessionId)
            }
            foreach ($p in $stillRunning) {
                try {
                    cmd.exe /c "taskkill /F /T /PID $($p.Id)" | Out-Null
                    Write-Log ("taskkill OK: {0} (PID {1})" -f $p.Name, $p.Id)
                    $killed++
                } catch {
                    Write-Log ("taskkill fehlgeschlagen für {0} (PID {1})" -f $p.Name,$p.Id) 'WARN'
                }
            }
        } else {
            Write-Log ("Keine passenden Prozesse in Session {0} gefunden." -f $sessionId)
        }

        # 4. Bis zu 15 Sekunden prüfen, ob noch etwas da ist
        $sw = [Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep 1
            $left = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                ($ProcessNames -contains $_.Name.ToLower()) -and ($_.SessionId -eq $sessionId)
            }
        } while ($left -and $sw.Elapsed.TotalSeconds -lt 15)

        if ($left) {
            Write-Log ("Laufen noch (Session {0}): {1}" -f $sessionId, ($left.Name -join ', ')) 'WARN'
        } else {
            Write-Log ("Alle Zielprozesse in Session {0} sind beendet." -f $sessionId)
        }

        Write-Log "Prozessliste nach Kill:"
        Get-Process | Where-Object { $_.Name -match 'teams|webview2|oneauth|tokenbroker|aad' } |
            Select-Object Name, Id, SessionId |
            ForEach-Object { Write-Log ("  {0} (PID {1}, Session {2})" -f $_.Name,$_.Id,$_.SessionId) }

    } catch {
        Write-Log ("Fehler beim Beenden von Prozessen: {0}" -f $_.Exception.Message) 'ERROR'
    }

    return $killed
}

# ==========================
# Sichere Pfad-Löschung (für Cache / Classic Teams)
# ==========================
function Remove-PathSafe {
    param([string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log ("Nicht vorhanden: {0}" -f $Path)
            return
        }

        cmd.exe /c "rd /s /q ""$Path""" | Out-Null

        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        }

        Write-Log ("Gelöscht: {0}" -f $Path)

    } catch {
        $msg = $_.Exception.Message

        if ($msg -like '*von einem anderen Prozess verwendet*') {
            Write-Log ("Pfad in Benutzung, übersprungen: {0}" -f $Path) 'WARN'
        }
        elseif ($msg -like '*Konflikt zwischen der Markierung*') {
            Write-Log ("ReparsePoint-Konflikt, übersprungen: {0}" -f $Path) 'WARN'
        }
        elseif ($msg -like '*konnte nicht gefunden werden*') {
            Write-Log ("Teilpfad schon entfernt: {0}" -f $Path) 'WARN'
        }
        else {
            Write-Log ("Fehler beim Löschen: {0} -> {1}" -f $Path, $msg) 'ERROR'
        }
    }
}

# ==========================
# Hard-Reset: Teams-AppX für Benutzer entfernen
# ==========================
function Invoke-TeamsAppxReset {
    Write-Log "Hard-Reset: Entferne Teams-AppX (MSTeams) für diesen Benutzer…"

    try {
        $pkg = Get-AppxPackage -Name 'MSTeams' -ErrorAction Stop
        Write-Log ("Gefundenes AppX-Paket: {0}" -f $pkg.PackageFullName)
    } catch {
        Write-Log "Teams-AppX für diesen Benutzer nicht gefunden – Hard-Reset übersprungen." 'WARN'
        return
    }

    try {
        Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop
        Write-Log "Teams-AppX erfolgreich entfernt. Teams wird beim nächsten Login automatisch neu bereitgestellt."
    } catch {
        Write-Log ("Remove-AppxPackage Fehler: {0}" -f $_.Exception.Message) 'ERROR'
    }

    # Classic-Teams-Dateipfade (falls vorhanden) mit aufräumen
    $classicDirs = @(
        (Join-Path $env:APPDATA      'Microsoft\Teams'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Teams')
    )
    foreach ($d in $classicDirs) {
        Remove-PathSafe -Path $d
    }

    # Classic-Teams-Registry (falls vorhanden) entfernen
    $regKeys = @(
        'HKCU:\Software\Microsoft\Office\Teams',
        'HKCU:\Software\Microsoft\Teams'
    )
    foreach ($r in $regKeys) {
        if (Test-Path $r) {
            try {
                Remove-Item $r -Recurse -Force -ErrorAction Stop
                Write-Log ("Registry gelöscht: {0}" -f $r)
            } catch {
                Write-Log ("Fehler beim Löschen der Registry {0}: {1}" -f $r,$_.Exception.Message) 'WARN'
            }
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
        [switch]$ClearCache,  # Teams 2 Cache / TempState
        [switch]$HardReset    # Remove-AppxPackage + Classic Cleanup
    )

    Write-Log "Starte Teams-Reset…"

    # 1. Prozesse beenden
    Stop-UserProcesses

    # 2. Hard-Reset?
    if ($HardReset) {
        Invoke-TeamsAppxReset
        Write-Log "Hard-Reset ausgeführt."
        return @{
            ExitCode = 0
            Errors   = 0
            Warnings = 0
        }
    }

    # 3. Normaler Reset: nur Cache löschen (wenn gewünscht)
    if ($ClearCache) {
        $pkgBase = Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe'
        $targets = @(
            (Join-Path $pkgBase 'LocalCache\Microsoft\MSTeams'),
            (Join-Path $pkgBase 'TempState')
        )

        foreach ($t in $targets) {
            Remove-PathSafe -Path $t
        }
        Write-Log "Teams-Cache-Reset abgeschlossen."
        } else {
            Write-Log "Kein Cache-Reset ausgewählt."
}

    return @{
        ExitCode = 0
        Errors   = 0
        Warnings = 0
    }
}

# ==========================
# GUI
# ==========================

$Form               = New-Object System.Windows.Forms.Form
$Form.Text          = "Microsoft Teams 2.0 – Reset"
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

Add-Label "Microsoft Teams 2.0 – Reset" 12 $true
Add-Label "Empfohlen: Teams-Cache löschen (schneller Standard-Reset bei Problemen)." 10 $false ([System.Drawing.Color]::DarkRed)
Add-Label "Hard-Reset: entfernt Teams vollständig für diesen Benutzer. Eine Neuanmeldung ist erforderlich, danach wird Teams neu bereitgestellt. Nur verwenden, wenn der Cache-Reset nicht hilft." 9 $false

# GroupBox für die 2 Optionen
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Aktionen"
$grp.Location = New-Object System.Drawing.Point(10, $y)
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

$cbCache = Add-Check "Teams-Cache löschen (empfohlen)"
$cbHard  = Add-Check "Hard-Reset: Teams-App entfernen (Neuanmeldung notwendig)"

# Standard: Cache-Reset aktiv, Hard-Reset aus
$cbCache.Checked = $true
$cbHard.Checked  = $false

# Gegenseitige Exklusivität: wenn Hard-Reset aktiv, Cache-Option ausgrauen
$cbHard.Add_CheckedChanged({
    if ($cbHard.Checked) {
        $cbCache.Checked = $false
        $cbCache.Enabled = $false
    } else {
        $cbCache.Enabled = $true
        $cbCache.Checked = $true   # Standard wiederherstellen
    }
})

$y += 140

# Buttons
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Reset ausführen"
$btnRun.Width = 150
$btnRun.Location = New-Object System.Drawing.Point(20, $y)
$Panel.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Schließen"
$btnClose.Width = 100
$btnClose.Location = New-Object System.Drawing.Point(190, $y)
$Panel.Controls.Add($btnClose)

$btnClose.Add_Click({ $Form.Close() })

$Form.AcceptButton = $btnRun
$Form.CancelButton = $btnClose

$btnRun.Add_Click({
    # Sicherheitscheck: falls durch irgendwas beide aus sind
    if (-not $cbCache.Checked -and -not $cbHard.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte wählen Sie mindestens eine Aktion aus (Cache-Reset oder Hard-Reset).",
            "Keine Aktion ausgewählt",
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
            "Fehler beim Ausführen des Resets:`n$($_.Exception.Message)",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $btnRun.Enabled = $true
    }
})

# GUI starten
[void]$Form.ShowDialog()
}




