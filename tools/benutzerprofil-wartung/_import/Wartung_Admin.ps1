<# ============================================================
    Wartung_Admin.ps1  (Management-Server / NETLOGON)
    ------------------------------------------------------------
    GUI-Editor für LogonOnce-Campaign (JSON)
    - Subskripte auswählen (aus NETLOGON\Scripte\*.ps1)
    - Parameter (enabled, campaignId, validUntil, cleanupDays, ...)
    - Schreibt Config\logon_once.json inkl. SHA256 pro Script
    - Optional: lädt bestehende JSON beim Start in die GUI
   ============================================================ #>

[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# 1) NETLOGON BASISPFAD (ANPASSEN WENN NÖTIG)
# =========================
$AdminToolRoot = "\\kontor-n.local\NETLOGON\Citrix\Scripte\Wartungsscript\v2\AdminTool"
$ScriptsRoot   = Join-Path $AdminToolRoot "Scripte"
$ConfigDir     = Join-Path $AdminToolRoot "Config"
$ConfigFile    = Join-Path $ConfigDir "logon_once.json"

# Default für centralLogRoot (kann leer bleiben)
$DefaultCentralLogRoot = ""

function Ensure-Dir([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) {
        New-Item -Path $p -ItemType Directory -Force | Out-Null
    }
}

function Get-Sha256Hex {
    param([Parameter(Mandatory)][string]$Path)
    (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Try-ParseDateYMD {
    param([string]$s)
    if (-not $s) { return $null }
    try { return [datetime]::ParseExact($s, "yyyy-MM-dd", $null) } catch { return $null }
}

function Read-ExistingConfig {
    if (-not (Test-Path -LiteralPath $ConfigFile)) { return $null }
    try {
        (Get-Content -LiteralPath $ConfigFile -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return $null
    }
}

Ensure-Dir $ConfigDir
if (-not (Test-Path -LiteralPath $ScriptsRoot)) {
    [System.Windows.Forms.MessageBox]::Show("ScriptsRoot nicht gefunden:`n$ScriptsRoot","Fehler",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit 1
}

# verfügbare Subskripte
$available = Get-ChildItem -Path $ScriptsRoot -Filter *.ps1 -File | Sort-Object Name

# =========================
# 2) GUI (FlowLayout: sauber & skalierbar)
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Wartung Admin – LogonOnce Campaign Builder"
$form.StartPosition = "CenterScreen"
$form.Width = 1200
$form.Height = 720
$form.MinimumSize = New-Object System.Drawing.Size(980,720)

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.ColumnCount = 1
$main.RowCount = 3
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($main)

# ---------- Header / Settings ----------
$settings = New-Object System.Windows.Forms.TableLayoutPanel
$settings.Dock = 'Top'
$settings.AutoSize = $true
$settings.ColumnCount = 4
$settings.RowCount = 4
$settings.Padding = '10,10,10,0'
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Wähle Subskripte für einmalige Logon-Ausführung (Campaign) aus und schreibe die JSON."
$lblTitle.AutoSize = $true
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblTitle.Margin = '0,0,0,10'
$settings.Controls.Add($lblTitle, 0, 0)
$settings.SetColumnSpan($lblTitle, 4)

# Row 1: campaignId + enabled
$lblCamp = New-Object System.Windows.Forms.Label
$lblCamp.Text = "CampaignId:"
$lblCamp.AutoSize = $true
$lblCamp.Anchor = 'Left'
$settings.Controls.Add($lblCamp, 0, 1)

$txtCamp = New-Object System.Windows.Forms.TextBox
$txtCamp.Anchor = 'Left,Right'
$txtCamp.Text = (Get-Date -Format "yyyy-MM-dd") + "_1"
$settings.Controls.Add($txtCamp, 1, 1)

$chkEnabled = New-Object System.Windows.Forms.CheckBox
$chkEnabled.Text = "enabled"
$chkEnabled.AutoSize = $true
$chkEnabled.Checked = $true
$chkEnabled.Anchor = 'Left'
$settings.Controls.Add($chkEnabled, 2, 1)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "Vorhandene JSON laden"
$btnLoad.AutoSize = $true
$btnLoad.Anchor = 'Right'
$settings.Controls.Add($btnLoad, 3, 1)

# Row 2: validUntil + centralLogRoot
$lblVU = New-Object System.Windows.Forms.Label
$lblVU.Text = "validUntil (yyyy-MM-dd):"
$lblVU.AutoSize = $true
$lblVU.Anchor = 'Left'
$settings.Controls.Add($lblVU, 0, 2)

$txtVU = New-Object System.Windows.Forms.TextBox
$txtVU.Anchor = 'Left,Right'
$txtVU.Text = (Get-Date).AddDays(4).ToString("yyyy-MM-dd")
$settings.Controls.Add($txtVU, 1, 2)

$lblCentral = New-Object System.Windows.Forms.Label
$lblCentral.Text = "centralLogRoot (optional UNC):"
$lblCentral.AutoSize = $true
$lblCentral.Anchor = 'Left'
$settings.Controls.Add($lblCentral, 2, 2)

$txtCentral = New-Object System.Windows.Forms.TextBox
$txtCentral.Anchor = 'Left,Right'
$txtCentral.Text = $DefaultCentralLogRoot
$settings.Controls.Add($txtCentral, 3, 2)

# Row 3: safety params (cleanupDays, maxRuntimeSeconds, shareTimeoutSeconds, stopOnFirstError)
$lblCleanup = New-Object System.Windows.Forms.Label
$lblCleanup.Text = "cleanupDays:"
$lblCleanup.AutoSize = $true
$lblCleanup.Anchor = 'Left'
$settings.Controls.Add($lblCleanup, 0, 3)

$txtCleanup = New-Object System.Windows.Forms.TextBox
$txtCleanup.Anchor = 'Left'
$txtCleanup.Width = 80
$txtCleanup.Text = "30"
$settings.Controls.Add($txtCleanup, 1, 3)

$panelSafetyRight = New-Object System.Windows.Forms.FlowLayoutPanel
$panelSafetyRight.FlowDirection = 'LeftToRight'
$panelSafetyRight.WrapContents = $false
$panelSafetyRight.AutoSize = $true
$panelSafetyRight.Anchor = 'Left'

$lblMaxRuntime = New-Object System.Windows.Forms.Label
$lblMaxRuntime.Text = "maxRuntimeSeconds:"
$lblMaxRuntime.AutoSize = $true
$lblMaxRuntime.Margin = '0,6,0,0'
$panelSafetyRight.Controls.Add($lblMaxRuntime)

$txtMaxRuntime = New-Object System.Windows.Forms.TextBox
$txtMaxRuntime.Width = 80
$txtMaxRuntime.Text = "600"
$panelSafetyRight.Controls.Add($txtMaxRuntime)

$lblShareTimeout = New-Object System.Windows.Forms.Label
$lblShareTimeout.Text = "shareTimeoutSeconds:"
$lblShareTimeout.AutoSize = $true
$lblShareTimeout.Margin = '10,6,0,0'
$panelSafetyRight.Controls.Add($lblShareTimeout)

$txtShareTimeout = New-Object System.Windows.Forms.TextBox
$txtShareTimeout.Width = 50
$txtShareTimeout.Text = "3"
$panelSafetyRight.Controls.Add($txtShareTimeout)

$chkStopOnErr = New-Object System.Windows.Forms.CheckBox
$chkStopOnErr.Text = "stopOnFirstError"
$chkStopOnErr.AutoSize = $true
$chkStopOnErr.Margin = '10,4,0,0'
$panelSafetyRight.Controls.Add($chkStopOnErr)

$settings.Controls.Add($panelSafetyRight, 2, 3)
$settings.SetColumnSpan($panelSafetyRight, 2)

$main.Controls.Add($settings, 0, 0)

# ---------- Grid ----------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.BackgroundColor = [System.Drawing.Color]::White

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.HeaderText = "Auswählen"
$colSel.FillWeight = 15
[void]$grid.Columns.Add($colSel)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.HeaderText = "Datei"
$colName.ReadOnly = $true
$colName.FillWeight = 40
[void]$grid.Columns.Add($colName)

$colArgs = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colArgs.HeaderText = "Args (optional)"
$colArgs.FillWeight = 45
[void]$grid.Columns.Add($colArgs)

foreach ($f in $available) {
    $i = $grid.Rows.Add()
    $grid.Rows[$i].Cells[0].Value = $false
    $grid.Rows[$i].Cells[1].Value = $f.Name
    $grid.Rows[$i].Cells[2].Value = ""
}

$main.Controls.Add($grid, 0, 1)

# ---------- Footer Buttons ----------
$footer = New-Object System.Windows.Forms.FlowLayoutPanel
$footer.Dock = 'Bottom'
$footer.AutoSize = $true
$footer.Padding = '10,10,10,10'
$footer.FlowDirection = 'LeftToRight'
$footer.WrapContents = $false

$btnWrite = New-Object System.Windows.Forms.Button
$btnWrite.Text = "JSON schreiben (enabled wie gesetzt)"
$btnWrite.AutoSize = $true

$btnDisable = New-Object System.Windows.Forms.Button
$btnDisable.Text = "Sofort deaktivieren (enabled=false)"
$btnDisable.AutoSize = $true

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "Config-Ordner öffnen"
$btnOpen.AutoSize = $true

$footer.Controls.AddRange(@($btnWrite, $btnDisable, $btnOpen))
$main.Controls.Add($footer, 0, 2)

$btnOpen.Add_Click({
    Start-Process explorer.exe "`"$ConfigDir`""
})

# =========================
# 3) LOGIK: GUI -> JSON
# =========================
function Get-IntOrDefault([string]$text, [int]$default) {
    try { return [int]($text.Trim()) } catch { return $default }
}

function Apply-ConfigToGui($cfg) {
    if (-not $cfg) { return }

    if ($cfg.campaignId) { $txtCamp.Text = [string]$cfg.campaignId }
    if ($null -ne $cfg.enabled) { $chkEnabled.Checked = [bool]$cfg.enabled }
    if ($cfg.validUntil) { $txtVU.Text = [string]$cfg.validUntil }
    if ($cfg.centralLogRoot) { $txtCentral.Text = [string]$cfg.centralLogRoot }

    # safety
    try { if ($cfg.safety.cleanupDays) { $txtCleanup.Text = [string][int]$cfg.safety.cleanupDays } } catch {}
    try { if ($cfg.safety.maxRuntimeSeconds) { $txtMaxRuntime.Text = [string][int]$cfg.safety.maxRuntimeSeconds } } catch {}
    try { if ($cfg.safety.shareTimeoutSeconds) { $txtShareTimeout.Text = [string][int]$cfg.safety.shareTimeoutSeconds } } catch {}
    try { if ($cfg.safety.stopOnFirstError) { $chkStopOnErr.Checked = [bool]$cfg.safety.stopOnFirstError } } catch {}

    # Subscripts: Haken setzen + args übernehmen, wenn Datei existiert
    if ($cfg.subScripts) {
        $map = @{}
        foreach ($s in $cfg.subScripts) {
            if ($s.file) { $map[[string]$s.file] = $s }
        }

        for ($r=0; $r -lt $grid.Rows.Count; $r++) {
            $file = [string]$grid.Rows[$r].Cells[1].Value
            if ($map.ContainsKey($file)) {
                $grid.Rows[$r].Cells[0].Value = $true
                try { $grid.Rows[$r].Cells[2].Value = [string]$map[$file].args } catch {}
            } else {
                $grid.Rows[$r].Cells[0].Value = $false
                $grid.Rows[$r].Cells[2].Value = ""
            }
        }
    }
}

function Build-And-WriteConfig([bool]$enabledOverride, [bool]$useOverride) {

    $campaignId = $txtCamp.Text.Trim()
    if (-not $campaignId) {
        [System.Windows.Forms.MessageBox]::Show("CampaignId fehlt.","Fehler") | Out-Null
        return
    }

    $vuText = $txtVU.Text.Trim()
    if ($vuText) {
        $vu = Try-ParseDateYMD $vuText
        if (-not $vu) {
            [System.Windows.Forms.MessageBox]::Show("validUntil ungültig. Erwartet yyyy-MM-dd.","Fehler") | Out-Null
            return
        }
    }

    $cleanupDays   = Get-IntOrDefault $txtCleanup.Text 30
    $maxRuntime    = Get-IntOrDefault $txtMaxRuntime.Text 600
    $shareTimeout  = Get-IntOrDefault $txtShareTimeout.Text 3
    $stopOnErr     = [bool]$chkStopOnErr.Checked

    if ($cleanupDays -lt 1 -or $cleanupDays -gt 365) {
        [System.Windows.Forms.MessageBox]::Show("cleanupDays muss zwischen 1 und 365 liegen.","Fehler") | Out-Null
        return
    }
    if ($maxRuntime -lt 30 -or $maxRuntime -gt 7200) {
        [System.Windows.Forms.MessageBox]::Show("maxRuntimeSeconds muss zwischen 30 und 7200 liegen.","Fehler") | Out-Null
        return
    }
    if ($shareTimeout -lt 1 -or $shareTimeout -gt 30) {
        [System.Windows.Forms.MessageBox]::Show("shareTimeoutSeconds muss zwischen 1 und 30 liegen.","Fehler") | Out-Null
        return
    }

    $enabled = if ($useOverride) { $enabledOverride } else { [bool]$chkEnabled.Checked }
    $central = $txtCentral.Text.Trim()

    $subs = New-Object System.Collections.ArrayList
    for ($r=0; $r -lt $grid.Rows.Count; $r++) {
        $sel = [bool]$grid.Rows[$r].Cells[0].Value
        if (-not $sel) { continue }

        $file = [string]$grid.Rows[$r].Cells[1].Value
        $args = [string]$grid.Rows[$r].Cells[2].Value

        $full = Join-Path $ScriptsRoot $file
        if (-not (Test-Path -LiteralPath $full)) {
            [System.Windows.Forms.MessageBox]::Show("Datei fehlt: $full","Fehler") | Out-Null
            return
        }

        $hash = Get-Sha256Hex -Path $full
        $name = [IO.Path]::GetFileNameWithoutExtension($file)

        [void]$subs.Add([pscustomobject]@{
            name   = $name
            file   = $file
            args   = $args
            sha256 = $hash
        })
    }

    if ($enabled -and $subs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("enabled=true aber keine Subskripte ausgewählt.","Hinweis") | Out-Null
        return
    }

    $cfg = [pscustomobject]@{
        enabled          = $enabled
        campaignId       = $campaignId
        validUntil       = $vuText
        scriptSourceRoot = $ScriptsRoot
        centralLogRoot   = $central
        subScripts       = $subs
        safety           = [pscustomobject]@{
            maxRuntimeSeconds    = $maxRuntime
            stopOnFirstError     = $stopOnErr
            shareTimeoutSeconds  = $shareTimeout
            cleanupDays          = $cleanupDays
        }
    }

    $json = $cfg | ConvertTo-Json -Depth 6

    Ensure-Dir $ConfigDir
    $tmp = Join-Path $ConfigDir ("logon_once_{0}.json.tmp" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    try {
        # "atomic-ish": erst tmp schreiben, dann replace
        Set-Content -LiteralPath $tmp -Value $json -Encoding UTF8 -Force
        Move-Item -LiteralPath $tmp -Destination $ConfigFile -Force

        [System.Windows.Forms.MessageBox]::Show(
            "OK: Config geschrieben:`n$ConfigFile`n`nenabled=$enabled; campaignId=$campaignId; scripts=$($subs.Count)`ncleanupDays=$cleanupDays",
            "Fertig",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Schreiben: $($_.Exception.Message)","Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        try { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# =========================
# 4) EVENTS
# =========================
$btnWrite.Add_Click({
    Build-And-WriteConfig -enabledOverride $false -useOverride:$false
})

$btnDisable.Add_Click({
    Build-And-WriteConfig -enabledOverride $false -useOverride:$true
})

$btnLoad.Add_Click({
    $cfg = Read-ExistingConfig
    if (-not $cfg) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine gültige bestehende JSON gefunden oder Parse-Fehler:`n$ConfigFile",
            "Hinweis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }
    Apply-ConfigToGui $cfg
    [System.Windows.Forms.MessageBox]::Show("Bestehende JSON geladen.","Info") | Out-Null
})

# Beim Start: bestehende Config automatisch laden (wenn vorhanden)
$existing = Read-ExistingConfig
if ($existing) { Apply-ConfigToGui $existing }

[void]$form.ShowDialog()