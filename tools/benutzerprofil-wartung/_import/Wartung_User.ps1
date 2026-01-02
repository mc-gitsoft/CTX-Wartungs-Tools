Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =====================================================================
# 0) KONFIGURATION: zentraler Admin-Logpfad
# =====================================================================

# UNC- oder lokaler Pfad, unter dem die Session-Logs zusaetzlich
# fuer Admins gesammelt werden sollen, z.B.:
# $AdminLogRoot = '\\kontor-n.local\NETLOGON\Citrix\WartungLogs'
# Leer lassen (''), wenn KEINE zentrale Kopie gewuenscht ist.
$AdminLogRoot = '\\fileserver\Userhomes'


# =====================================================================
# 1) AKTIONEN DEFINIEREN (Reihenfolge = Anzeige-Reihenfolge)
# =====================================================================

$Aktionen = @(
    @{
        Id             = 'Teams_Reset'
        Name           = 'Teams zuruecksetzen'
        Skript         = 'Teams_Reset.ps1'
        Kategorie      = 'Kommunikation'
        ConfirmMessage = ''
        Beschreibung   = 'Bereinigt Microsoft Teams Cache und Konfigurationsdateien im Benutzerprofil.'
        Interaktiv     = $true
    },
    @{
        Id             = 'WorkspaceApp_Reset'
        Name           = 'Citrix Workspace App zuruecksetzen'
        Skript         = 'WorkspaceApp_Reset.ps1'
        Kategorie      = 'Citrix'
        ConfirmMessage = 'Citrix Workspace App Einstellungen werden geloescht. Fortfahren?'
        Beschreibung   = 'Setzt Citrix Workspace-App-Einstellungen zurueck.'
        Interaktiv     = $true
    },
    @{
        Id             = 'Edge_Reset'
        Name           = 'Microsoft Edge Browser zuruecksetzen'
        Skript         = 'MS_Edge_Reset.ps1'
        Kategorie      = 'Browser'
        ConfirmMessage = 'Edge-Einstellungen und lokale Browserdaten werden geloescht. Fortfahren?'
        Beschreibung   = 'Bereinigt Edge Cache, Profile und Konfigurationen.'
        Interaktiv     = $true
    },
    @{
        Id             = 'Chrome_Reset'
        Name           = 'Google Chrome Browser zuruecksetzen'
        Skript         = 'Google_Chrome_Reset.ps1'
        Kategorie      = 'Browser'
        ConfirmMessage = 'Chrome-Einstellungen und lokale Browserdaten werden geloescht. Fortfahren?'
        Beschreibung   = 'Bereinigt Chrome Cache, Profile und Konfigurationen.'
        Interaktiv     = $true
    },
    @{
        Id             = 'AdobeReader_Reset'
        Name           = 'Adobe Reader zuruecksetzen'
        Skript         = 'AdobeReader_Reset.ps1'
        Kategorie      = 'PDF'
        ConfirmMessage = 'Adobe Reader Einstellungen werden geloescht. Fortfahren?'
        Beschreibung   = 'Setzt Adobe Reader Profil und Konfigurationen zurueck.'
        Interaktiv     = $true
    },
    @{
        Id             = 'Outlook_Reset'
        Name           = 'Microsoft Outlook zuruecksetzen'
        Skript         = 'Outlook_Reset.ps1'
        Kategorie      = 'Office'
        ConfirmMessage = ''
        Beschreibung   = 'Bereinigt Outlook Cache, Registry-Reste und Suchindex (Benutzerkontext).'
        Interaktiv     = $true
    },
    @{
        Id             = 'ProfileData_Import'
        Name           = 'Profildaten-Import'
        Skript         = '\\kontor-n.local\NETLOGON\Citrix\Scripte\Profilmigration_FSLogix\Profildaten-Import.ps1'
        Kategorie      = 'Profile'
        ConfirmMessage = ''
        Beschreibung   = 'Importiert Benutzerprofile (Favoriten, Signaturen, etc.).'
        Interaktiv     = $true
    }
)

# =====================================================================
# 2) PFAD STRUKTUR
# =====================================================================

$ScriptPath  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ScriptePath = Join-Path -Path $ScriptPath -ChildPath 'Scripte'

# =====================================================================
# 3) LOGGING + SESSION ORDNER (lokal im Benutzerprofil)
# =====================================================================

$BaseLogRoot = Join-Path $env:LOCALAPPDATA 'Wartungsscript\Logs'
if (-not (Test-Path $BaseLogRoot)) {
    New-Item -Path $BaseLogRoot -ItemType Directory -Force | Out-Null
}

$SessionTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SessionFolder = Join-Path $BaseLogRoot ("Log_{0}" -f $SessionTimestamp)
New-Item -Path $SessionFolder -ItemType Directory -Force | Out-Null

$LogFile = Join-Path $SessionFolder 'Wartung_User.log'

# --- Log Rotation (max 20 Sessions) ---
$allSessions = Get-ChildItem -Path $BaseLogRoot -Directory | Sort-Object CreationTime -Descending
if ($allSessions.Count -gt 20) {
    $allSessions | Select-Object -Skip 20 | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $user      = $env:USERNAME
    Add-Content -Path $LogFile -Value "$timestamp;$Level;$user;$Message"
}

Write-Host "Session-Logs unter: $SessionFolder" -ForegroundColor DarkGray
Write-Log "Wartungstool gestartet. ScriptPath=$ScriptPath; ScriptePath=$ScriptePath"


# =====================================================================
# 4) GUI AUFBAU
# =====================================================================

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Wartung Ausfuehren (Benutzer)"
$Form.StartPosition = "CenterScreen"
$Form.AutoSize = $true
$Form.AutoSizeMode = "GrowAndShrink"
$Form.MaximizeBox = $false

$Flow = New-Object System.Windows.Forms.FlowLayoutPanel
$Flow.Dock = 'Fill'
$Flow.FlowDirection = 'TopDown'
$Flow.WrapContents = $false
$Flow.AutoSize = $true

# --- Label ---
$Label = New-Object System.Windows.Forms.Label
$Label.AutoSize = $true
$Label.Text = "Aktion(en) auswaehlen:"
$Flow.Controls.Add($Label)

# --- Checkboxes ---
$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.AutoPopDelay = 10000

$CheckboxMap = @{}

foreach ($aktion in $Aktionen) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $aktion.Name
    $cb.AutoSize = $true
    $cb.Tag = $aktion

    if ($aktion.Beschreibung) {
        $ToolTip.SetToolTip($cb, $aktion.Beschreibung)
    }

    $Flow.Controls.Add($cb)
    $CheckboxMap[$aktion.Id] = $cb
}

# --- Buttons ---
$Button = New-Object System.Windows.Forms.Button
$Button.Text = "Ausfuehren"
$Button.AutoSize = $true

$LogButton = New-Object System.Windows.Forms.Button
$LogButton.Text = "Log-Ordner oeffnen"
$LogButton.AutoSize = $true

$ButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$ButtonPanel.FlowDirection = 'LeftToRight'
$ButtonPanel.WrapContents = $false
$ButtonPanel.AutoSize = $true
$ButtonPanel.Margin = '0,10,0,0'

$ButtonPanel.Controls.Add($Button)
$ButtonPanel.Controls.Add($LogButton)

$Flow.Controls.Add($ButtonPanel)

# =====================================================================
# 5) AKTIONEN AUSFUEHREN
# =====================================================================

function Start-UserActions {
    param(
        $CheckboxMap, $Aktionen, $ScriptePath, $SessionFolder
    )

    $mindestensEine = $false

    foreach ($aktion in $Aktionen) {
        $cb = $CheckboxMap[$aktion.Id]
        if (-not $cb.Checked) { continue }

        Write-Log "Aktion gewaehlt: $($aktion.Id)"

        # --- Confirm ---
        if ($aktion.ConfirmMessage) {
            $res = [System.Windows.Forms.MessageBox]::Show(
                $aktion.ConfirmMessage, $aktion.Name,
                'YesNo', 'Warning'
            )
            if ($res -ne 'Yes') {
                Write-Log "Abgebrochen ueber Confirm: $($aktion.Id)" 'WARN'
                continue
            }
        }

        # --- Pfad ---
        $skriptPfad =
            if ([IO.Path]::IsPathRooted($aktion.Skript) -or $aktion.Skript.StartsWith('\')) {
                $aktion.Skript
            } else {
                Join-Path $ScriptePath $aktion.Skript
            }

        if (-not (Test-Path $skriptPfad)) {
            Write-Log "FEHLER: Skript nicht gefunden: $skriptPfad" 'ERROR'
            Continue
        }

        $subLog = Join-Path $SessionFolder ("{0}.log" -f $aktion.Id)

        if ($aktion.Interaktiv) {
            $command = "Start-Transcript -Path '$subLog' -Force; & '$skriptPfad'; Stop-Transcript"

            $proc = Start-Process powershell.exe `
                -ArgumentList "-NoLogo -NoProfile -ExecutionPolicy Bypass -Command $command" `
                -WorkingDirectory (Split-Path $skriptPfad) `
                -WindowStyle Normal `
                -PassThru

            $proc.WaitForExit()   # <- BLOCKIERT den GUI-Thread
        }

        else {
            # --- Nicht interaktiv: Ausgabe abfangen ---
            Write-Log "Starte: $skriptPfad; Log=$subLog"

            $psArgs = @(
                '-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-File',$skriptPfad
            )

            $output = & powershell.exe @psArgs 2>&1
            $exit = $LASTEXITCODE

            Add-Content -Path $subLog -Value "### START ###"
            foreach ($line in $output) {
                Add-Content -Path $subLog -Value ("OUT: " + $line.ToString())
            }
            Add-Content -Path $subLog -Value "EXITCODE=$exit"
            Add-Content -Path $subLog -Value "### ENDE ###"

            Write-Log "Beendet: $skriptPfad Exit=$exit"
        }

        $mindestensEine = $true
    }

    return $mindestensEine
}

# =====================================================================
# 6) BUTTON EVENTS
# =====================================================================

$Button.Add_Click({
    if (Start-UserActions -CheckboxMap $CheckboxMap -Aktionen $Aktionen -ScriptePath $ScriptePath -SessionFolder $SessionFolder) {
        Write-Log "Aktionen abgeschlossen."
        [System.Windows.Forms.MessageBox]::Show("Aktionen abgeschlossen.", "Fertig", 'OK', 'Information')
        foreach ($cb in $CheckboxMap.Values) { $cb.Checked = $false }
    }
    else {
        Write-Log "Keine Aktion ausgewaehlt." 'WARN'
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Aktion auswaehlen.", "Hinweis", 'OK', 'Warning')
    }
})

$LogButton.Add_Click({
    Start-Process explorer.exe "`"$SessionFolder`""
})

# =====================================================================
# 7) START GUI
# =====================================================================

$Form.Controls.Add($Flow)
$Form.ShowDialog() | Out-Null
Write-Log "GUI geschlossen."

# =====================================================================
# 8) Session-Logs optional in zentralen Admin-Pfad kopieren + Rotation
# =====================================================================

if ($AdminLogRoot -and $AdminLogRoot.Trim().Length -gt 0) {
    try {
        if (Test-Path $AdminLogRoot) {

            # Ziel: <AdminLogRoot>\<USERNAME>\Wartungsskript\Log_YYYYMMDD_HHMMSS
            $UserRoot            = Join-Path $AdminLogRoot $env:USERNAME
            $UserWartungRoot     = Join-Path $UserRoot 'Wartungsskript'
            $targetSessionFolder = Join-Path $UserWartungRoot (Split-Path $SessionFolder -Leaf)

            # Ordnerstruktur erzeugen
            if (-not (Test-Path $UserRoot)) {
                New-Item -Path $UserRoot -ItemType Directory -Force | Out-Null
            }
            if (-not (Test-Path $UserWartungRoot)) {
                New-Item -Path $UserWartungRoot -ItemType Directory -Force | Out-Null
            }

            # --- Session kopieren ---
            Copy-Item -Path $SessionFolder -Destination $targetSessionFolder -Recurse -Force -ErrorAction Stop

            Write-Log "Session-Log nach zentralem Admin-Pfad kopiert: $targetSessionFolder"

            # =====================================================================
            # Admin-Log-Rotation (max. 20 Sessions behalten)
            # =====================================================================

            try {
                $adminSessions = Get-ChildItem -Path $UserWartungRoot -Directory |
                                 Sort-Object CreationTime -Descending

                if ($adminSessions.Count -gt 20) {
                    $delete = $adminSessions | Select-Object -Skip 20
                    foreach ($old in $delete) {
                        try {
                            Remove-Item -Path $old.FullName -Recurse -Force -ErrorAction Stop
                            Write-Log "Alten Admin-Log entfernt: $($old.FullName)"
                        }
                        catch {
                            Write-Log "Fehler beim Entfernen alter Admin-Logs: $($_.Exception.Message)" 'WARN'
                        }
                    }
                }
            }
            catch {
                Write-Log "Admin-Log-Rotation fehlgeschlagen: $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Log "AdminLogRoot '$AdminLogRoot' nicht erreichbar - zentrale Kopie uebersprungen." 'WARN'
        }
    }
    catch {
        Write-Log "Fehler beim Kopieren der Session-Logs in Admin-Pfad '$AdminLogRoot': $($_.Exception.Message)" 'WARN'
    }
}
else {
    Write-Log "AdminLogRoot nicht gesetzt - keine zentrale Logkopie." 'INFO'
}
