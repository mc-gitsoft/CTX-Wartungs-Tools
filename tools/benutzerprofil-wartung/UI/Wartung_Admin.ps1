[CmdletBinding()]
param()

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$modulePath = Join-Path $toolRoot "shared\WartungsTools.SDK.psm1"
Import-Module $modulePath -Force

$toolManifest = Get-Content (Join-Path $toolRoot "tool.json") -Raw | ConvertFrom-Json
$toolId = $toolManifest.toolId

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$actionsPath = Join-Path $toolRoot "Actions"
$available = Get-ChildItem -Path $actionsPath -Filter *.ps1 -File | Sort-Object Name

function Parse-ParamsText {
    param([string]$Text)

    if (-not $Text -or -not $Text.Trim()) { return @{} }

    try {
        $obj = $Text | ConvertFrom-Json -ErrorAction Stop
        return $obj
    } catch {
        $pairs = $Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $ht = @{}
        foreach ($pair in $pairs) {
            $kv = $pair -split '=', 2
            if ($kv.Count -eq 2) {
                $ht[$kv[0].Trim()] = $kv[1].Trim()
            }
        }
        return $ht
    }
}

function Read-ExistingPolicy {
    $path = Join-Path $toolRoot "policy.json"
    if (-not (Test-Path $path)) { return $null }
    try {
        return (Get-Content $path -Raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Get-CustomerPath {
    return (Join-Path $toolRoot "customer.json")
}

function Read-ExistingCustomer {
    $path = Get-CustomerPath
    if (-not (Test-Path $path)) { return $null }
    try {
        return (Get-Content $path -Raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Is-RelativePath {
    param([string]$Path)
    if (-not $Path) { return $true }
    if ($Path.StartsWith("\\\\")) { return $false }
    return -not [IO.Path]::IsPathRooted($Path)
}

function New-DefaultCustomerConfig {
    return [pscustomobject]@{
        customer = [pscustomobject]@{
            name = ""
        }
        paths = [pscustomobject]@{
            repoRoot = ""
        }
        fslogix = [pscustomobject]@{
            enabled = $false
            profileShare = ""
            officeContainerShare = ""
        }
        branding = [pscustomobject]@{
            windowTitle = "CTX Wartung"
            supportText = ""
        }
        logging = [pscustomobject]@{
            relativeLogPath = "logs"
            adminLogRoot = ""
        }
        flags = [pscustomobject]@{
            allowOffline = $true
            allowLogoffRunner = $true
        }
    }
}

function Validate-CustomerConfig {
    param([pscustomobject]$Config)

    if (-not $Config.customer -or -not $Config.customer.name -or -not $Config.customer.name.Trim()) {
        return "customer.name darf nicht leer sein."
    }
    if (-not $Config.branding -or -not $Config.branding.windowTitle -or -not $Config.branding.windowTitle.Trim()) {
        return "branding.windowTitle darf nicht leer sein."
    }
    if (-not $Config.logging -or -not $Config.logging.relativeLogPath -or -not $Config.logging.relativeLogPath.Trim()) {
        return "logging.relativeLogPath darf nicht leer sein."
    }
    if (-not (Is-RelativePath -Path $Config.logging.relativeLogPath)) {
        return "logging.relativeLogPath darf kein absoluter Pfad sein."
    }

    return $null
}

$form = New-Object System.Windows.Forms.Form
$baseTitle = "Wartung Admin"
$form.Text = $baseTitle
$form.StartPosition = "CenterScreen"
$form.Width = 1100
$form.Height = 700
$form.MinimumSize = New-Object System.Drawing.Size(900,650)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'

$tabPolicy = New-Object System.Windows.Forms.TabPage
$tabPolicy.Text = "Policy"

$tabConfig = New-Object System.Windows.Forms.TabPage
$tabConfig.Text = "Konfiguration"

[void]$tabs.TabPages.AddRange(@($tabPolicy, $tabConfig))
$form.Controls.Add($tabs)

# =========================
# Policy Tab
# =========================

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.ColumnCount = 1
$main.RowCount = 3
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tabPolicy.Controls.Add($main)

$settings = New-Object System.Windows.Forms.TableLayoutPanel
$settings.Dock = 'Top'
$settings.AutoSize = $true
$settings.ColumnCount = 4
$settings.RowCount = 3
$settings.Padding = '10,10,10,0'
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$settings.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Logon-Once Policy erstellen (policy.json im Tool-Root)."
$lblTitle.AutoSize = $true
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblTitle.Margin = '0,0,0,10'
$settings.Controls.Add($lblTitle, 0, 0)
$settings.SetColumnSpan($lblTitle, 4)

$lblCamp = New-Object System.Windows.Forms.Label
$lblCamp.Text = "CampaignId:"
$lblCamp.AutoSize = $true
$lblCamp.Anchor = 'Left'
$settings.Controls.Add($lblCamp, 0, 1)

$txtCamp = New-Object System.Windows.Forms.TextBox
$txtCamp.Anchor = 'Left,Right'
$txtCamp.Text = (Get-Date -Format "yyyy-MM-dd") + "_1"
$settings.Controls.Add($txtCamp, 1, 1)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "Vorhandene Policy laden"
$btnLoad.AutoSize = $true
$btnLoad.Anchor = 'Right'
$settings.Controls.Add($btnLoad, 3, 1)

$lblVU = New-Object System.Windows.Forms.Label
$lblVU.Text = "validUntil (yyyy-MM-dd):"
$lblVU.AutoSize = $true
$lblVU.Anchor = 'Left'
$settings.Controls.Add($lblVU, 0, 2)

$txtVU = New-Object System.Windows.Forms.TextBox
$txtVU.Anchor = 'Left,Right'
$txtVU.Text = (Get-Date).AddDays(4).ToString("yyyy-MM-dd")
$settings.Controls.Add($txtVU, 1, 2)

$lblUsers = New-Object System.Windows.Forms.Label
$lblUsers.Text = "Targets users (comma):"
$lblUsers.AutoSize = $true
$lblUsers.Anchor = 'Left'
$settings.Controls.Add($lblUsers, 2, 2)

$txtUsers = New-Object System.Windows.Forms.TextBox
$txtUsers.Anchor = 'Left,Right'
$txtUsers.Text = ""
$settings.Controls.Add($txtUsers, 3, 2)

$main.Controls.Add($settings, 0, 0)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.BackgroundColor = [System.Drawing.Color]::White

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.HeaderText = "Auswaehlen"
$colSel.FillWeight = 15
[void]$grid.Columns.Add($colSel)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.HeaderText = "Action"
$colName.ReadOnly = $true
$colName.FillWeight = 40
[void]$grid.Columns.Add($colName)

$colParams = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colParams.HeaderText = "Params (JSON oder key=value;key2=value2)"
$colParams.FillWeight = 45
[void]$grid.Columns.Add($colParams)

foreach ($f in $available) {
    $i = $grid.Rows.Add()
    $grid.Rows[$i].Cells[0].Value = $false
    $grid.Rows[$i].Cells[1].Value = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $grid.Rows[$i].Cells[2].Value = ""
}

$main.Controls.Add($grid, 0, 1)

$footer = New-Object System.Windows.Forms.FlowLayoutPanel
$footer.Dock = 'Bottom'
$footer.AutoSize = $true
$footer.Padding = '10,10,10,10'
$footer.FlowDirection = 'LeftToRight'
$footer.WrapContents = $false

$btnWrite = New-Object System.Windows.Forms.Button
$btnWrite.Text = "Policy schreiben (logon.once)"
$btnWrite.AutoSize = $true

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "Policy-Ordner oeffnen"
$btnOpen.AutoSize = $true

$footer.Controls.AddRange(@($btnWrite, $btnOpen))
$main.Controls.Add($footer, 0, 2)

$btnOpen.Add_Click({
    Start-Process explorer.exe "`"$toolRoot`""
})

function Apply-PolicyToGui($policy) {
    if (-not $policy -or -not $policy.logon -or -not $policy.logon.once) { return }

    $entry = $policy.logon.once | Select-Object -First 1
    if ($entry.campaignId) { $txtCamp.Text = [string]$entry.campaignId }
    if ($entry.validUntil) { $txtVU.Text = [string]$entry.validUntil }
    if ($entry.targets -and $entry.targets.users) {
        $txtUsers.Text = ($entry.targets.users -join ',')
    }

    $actionMap = @{}
    foreach ($a in $entry.actions) {
        if ($a.name) { $actionMap[[string]$a.name] = $a }
    }

    for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
        $name = [string]$grid.Rows[$r].Cells[1].Value
        if ($actionMap.ContainsKey($name)) {
            $grid.Rows[$r].Cells[0].Value = $true
            $params = $actionMap[$name].params
            if ($params) {
                $grid.Rows[$r].Cells[2].Value = ($params | ConvertTo-Json -Compress)
            }
        } else {
            $grid.Rows[$r].Cells[0].Value = $false
            $grid.Rows[$r].Cells[2].Value = ""
        }
    }
}

$btnLoad.Add_Click({
    $policy = Read-ExistingPolicy
    if (-not $policy) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine gueltige policy.json gefunden.",
            "Hinweis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }
    Apply-PolicyToGui $policy
    [System.Windows.Forms.MessageBox]::Show("Policy geladen.", "Info") | Out-Null
})

$btnWrite.Add_Click({
    $campaignId = $txtCamp.Text.Trim()
    if (-not $campaignId) {
        [System.Windows.Forms.MessageBox]::Show("CampaignId fehlt.", "Fehler") | Out-Null
        return
    }

    $validUntil = $txtVU.Text.Trim()
    if ($validUntil) {
        try { [datetime]::ParseExact($validUntil, "yyyy-MM-dd", $null) | Out-Null }
        catch {
            [System.Windows.Forms.MessageBox]::Show("validUntil ungueltig. Erwartet yyyy-MM-dd.", "Fehler") | Out-Null
            return
        }
    }

    $users = $txtUsers.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    $actionsList = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
        $sel = [bool]$grid.Rows[$r].Cells[0].Value
        if (-not $sel) { continue }

        $name = [string]$grid.Rows[$r].Cells[1].Value
        $paramsText = [string]$grid.Rows[$r].Cells[2].Value
        $params = Parse-ParamsText $paramsText

        [void]$actionsList.Add([pscustomobject]@{
            name = $name
            params = $params
            mode = "Silent"
        })
    }

    if ($actionsList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Action auswaehlen.", "Hinweis") | Out-Null
        return
    }

    $entry = [pscustomobject]@{
        campaignId = $campaignId
        validUntil = $validUntil
        actions = $actionsList
    }

    if ($users.Count -gt 0) {
        $entry | Add-Member -NotePropertyName targets -NotePropertyValue ([pscustomobject]@{ users = $users })
    }

    $existing = Read-ExistingPolicy
    if (-not $existing) {
        $existing = [pscustomobject]@{
            logon = [pscustomobject]@{ every = @(); once = @() }
            logoff = [pscustomobject]@{ every = @(); once = @() }
        }
    }

    $existing.logon.once = @($entry)

    $json = $existing | ConvertTo-Json -Depth 6
    $policyPath = Join-Path $toolRoot "policy.json"
    Set-Content -Path $policyPath -Value $json -Encoding UTF8

    [System.Windows.Forms.MessageBox]::Show("Policy geschrieben: $policyPath", "Fertig") | Out-Null
})

# =========================
# Konfiguration Tab
# =========================

$configMain = New-Object System.Windows.Forms.TableLayoutPanel
$configMain.Dock = 'Fill'
$configMain.ColumnCount = 2
$configMain.RowCount = 13
$configMain.Padding = '10,10,10,10'
$configMain.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$configMain.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabConfig.Controls.Add($configMain)

$lblCfgTitle = New-Object System.Windows.Forms.Label
$lblCfgTitle.Text = "customer.json im Tool-Root erstellen/bearbeiten."
$lblCfgTitle.AutoSize = $true
$lblCfgTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblCfgTitle.Margin = '0,0,0,10'
$configMain.Controls.Add($lblCfgTitle, 0, 0)
$configMain.SetColumnSpan($lblCfgTitle, 2)

$lblCfgWarning = New-Object System.Windows.Forms.Label
$lblCfgWarning.Text = ""
$lblCfgWarning.AutoSize = $true
$lblCfgWarning.ForeColor = [System.Drawing.Color]::DarkRed
$lblCfgWarning.Margin = '0,0,0,10'
$lblCfgWarning.Visible = $false
$configMain.Controls.Add($lblCfgWarning, 0, 1)
$configMain.SetColumnSpan($lblCfgWarning, 2)

$lblCustomerName = New-Object System.Windows.Forms.Label
$lblCustomerName.Text = "Customer Name*"
$lblCustomerName.AutoSize = $true
$configMain.Controls.Add($lblCustomerName, 0, 2)

$txtCustomerName = New-Object System.Windows.Forms.TextBox
$txtCustomerName.Anchor = 'Left,Right'
$configMain.Controls.Add($txtCustomerName, 1, 2)

$lblRepoRoot = New-Object System.Windows.Forms.Label
$lblRepoRoot.Text = "RepoRoot (optional)"
$lblRepoRoot.AutoSize = $true
$configMain.Controls.Add($lblRepoRoot, 0, 3)

$txtRepoRoot = New-Object System.Windows.Forms.TextBox
$txtRepoRoot.Anchor = 'Left,Right'
$configMain.Controls.Add($txtRepoRoot, 1, 3)

$lblFslogixEnabled = New-Object System.Windows.Forms.Label
$lblFslogixEnabled.Text = "FSLogix enabled"
$lblFslogixEnabled.AutoSize = $true
$configMain.Controls.Add($lblFslogixEnabled, 0, 4)

$chkFslogixEnabled = New-Object System.Windows.Forms.CheckBox
$chkFslogixEnabled.Text = ""
$chkFslogixEnabled.AutoSize = $true
$configMain.Controls.Add($chkFslogixEnabled, 1, 4)

$lblProfileShare = New-Object System.Windows.Forms.Label
$lblProfileShare.Text = "FSLogix profileShare (optional)"
$lblProfileShare.AutoSize = $true
$configMain.Controls.Add($lblProfileShare, 0, 5)

$txtProfileShare = New-Object System.Windows.Forms.TextBox
$txtProfileShare.Anchor = 'Left,Right'
$configMain.Controls.Add($txtProfileShare, 1, 5)

$lblOfficeShare = New-Object System.Windows.Forms.Label
$lblOfficeShare.Text = "FSLogix officeContainerShare (optional)"
$lblOfficeShare.AutoSize = $true
$configMain.Controls.Add($lblOfficeShare, 0, 6)

$txtOfficeShare = New-Object System.Windows.Forms.TextBox
$txtOfficeShare.Anchor = 'Left,Right'
$configMain.Controls.Add($txtOfficeShare, 1, 6)

$lblWindowTitle = New-Object System.Windows.Forms.Label
$lblWindowTitle.Text = "WindowTitle*"
$lblWindowTitle.AutoSize = $true
$configMain.Controls.Add($lblWindowTitle, 0, 7)

$txtWindowTitle = New-Object System.Windows.Forms.TextBox
$txtWindowTitle.Anchor = 'Left,Right'
$configMain.Controls.Add($txtWindowTitle, 1, 7)

$lblSupportText = New-Object System.Windows.Forms.Label
$lblSupportText.Text = "SupportText (optional)"
$lblSupportText.AutoSize = $true
$configMain.Controls.Add($lblSupportText, 0, 8)

$txtSupportText = New-Object System.Windows.Forms.TextBox
$txtSupportText.Anchor = 'Left,Right'
$configMain.Controls.Add($txtSupportText, 1, 8)

$lblLogPath = New-Object System.Windows.Forms.Label
$lblLogPath.Text = "relativeLogPath*"
$lblLogPath.AutoSize = $true
$configMain.Controls.Add($lblLogPath, 0, 9)

$txtRelativeLogPath = New-Object System.Windows.Forms.TextBox
$txtRelativeLogPath.Anchor = 'Left,Right'
$configMain.Controls.Add($txtRelativeLogPath, 1, 9)

$lblAdminLog = New-Object System.Windows.Forms.Label
$lblAdminLog.Text = "adminLogRoot (optional)"
$lblAdminLog.AutoSize = $true
$configMain.Controls.Add($lblAdminLog, 0, 10)

$txtAdminLogRoot = New-Object System.Windows.Forms.TextBox
$txtAdminLogRoot.Anchor = 'Left,Right'
$configMain.Controls.Add($txtAdminLogRoot, 1, 10)

$lblFlags = New-Object System.Windows.Forms.Label
$lblFlags.Text = "Flags"
$lblFlags.AutoSize = $true
$configMain.Controls.Add($lblFlags, 0, 11)

$flagsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flagsPanel.FlowDirection = 'LeftToRight'
$flagsPanel.AutoSize = $true
$flagsPanel.WrapContents = $false

$chkAllowOffline = New-Object System.Windows.Forms.CheckBox
$chkAllowOffline.Text = "allowOffline"
$chkAllowOffline.AutoSize = $true
$flagsPanel.Controls.Add($chkAllowOffline)

$chkAllowLogoffRunner = New-Object System.Windows.Forms.CheckBox
$chkAllowLogoffRunner.Text = "allowLogoffRunner"
$chkAllowLogoffRunner.AutoSize = $true
$flagsPanel.Controls.Add($chkAllowLogoffRunner)

$configMain.Controls.Add($flagsPanel, 1, 11)

$cfgFooter = New-Object System.Windows.Forms.FlowLayoutPanel
$cfgFooter.FlowDirection = 'LeftToRight'
$cfgFooter.AutoSize = $true
$cfgFooter.WrapContents = $false

$btnSaveCustomer = New-Object System.Windows.Forms.Button
$btnSaveCustomer.Text = "customer.json speichern"
$btnSaveCustomer.AutoSize = $true

$btnNewCustomer = New-Object System.Windows.Forms.Button
$btnNewCustomer.Text = "Neu anlegen (Defaults)"
$btnNewCustomer.AutoSize = $true

$btnOpenCustomer = New-Object System.Windows.Forms.Button
$btnOpenCustomer.Text = "customer.json oeffnen"
$btnOpenCustomer.AutoSize = $true

$btnOpenLogs = New-Object System.Windows.Forms.Button
$btnOpenLogs.Text = "Logs-Ordner oeffnen"
$btnOpenLogs.AutoSize = $true

$cfgFooter.Controls.AddRange(@($btnSaveCustomer, $btnNewCustomer, $btnOpenCustomer, $btnOpenLogs))
$configMain.Controls.Add($cfgFooter, 1, 12)

function Apply-CustomerToGui {
    param([pscustomobject]$Config)

    $txtCustomerName.Text = [string]$Config.customer.name
    $txtRepoRoot.Text = [string]$Config.paths.repoRoot
    $chkFslogixEnabled.Checked = [bool]$Config.fslogix.enabled
    $txtProfileShare.Text = [string]$Config.fslogix.profileShare
    $txtOfficeShare.Text = [string]$Config.fslogix.officeContainerShare
    $txtWindowTitle.Text = [string]$Config.branding.windowTitle
    $txtSupportText.Text = [string]$Config.branding.supportText
    $txtRelativeLogPath.Text = [string]$Config.logging.relativeLogPath
    $txtAdminLogRoot.Text = [string]$Config.logging.adminLogRoot
    $chkAllowOffline.Checked = [bool]$Config.flags.allowOffline
    $chkAllowLogoffRunner.Checked = [bool]$Config.flags.allowLogoffRunner
}

function Build-CustomerFromGui {
    return [pscustomobject]@{
        customer = [pscustomobject]@{
            name = $txtCustomerName.Text.Trim()
        }
        paths = [pscustomobject]@{
            repoRoot = $txtRepoRoot.Text.Trim()
        }
        fslogix = [pscustomobject]@{
            enabled = [bool]$chkFslogixEnabled.Checked
            profileShare = $txtProfileShare.Text.Trim()
            officeContainerShare = $txtOfficeShare.Text.Trim()
        }
        branding = [pscustomobject]@{
            windowTitle = $txtWindowTitle.Text.Trim()
            supportText = $txtSupportText.Text.Trim()
        }
        logging = [pscustomobject]@{
            relativeLogPath = $txtRelativeLogPath.Text.Trim()
            adminLogRoot = $txtAdminLogRoot.Text.Trim()
        }
        flags = [pscustomobject]@{
            allowOffline = [bool]$chkAllowOffline.Checked
            allowLogoffRunner = [bool]$chkAllowLogoffRunner.Checked
        }
    }
}

$btnSaveCustomer.Add_Click({
    $cfg = Build-CustomerFromGui
    $validationError = Validate-CustomerConfig -Config $cfg
    if ($validationError) {
        [System.Windows.Forms.MessageBox]::Show($validationError, "Fehler") | Out-Null
        return
    }

    $json = $cfg | ConvertTo-Json -Depth 5
    $path = Get-CustomerPath
    Set-Content -Path $path -Value $json -Encoding UTF8
    $btnNewCustomer.Visible = $false
    $lblCfgWarning.Visible = $false
    $form.Text = ("{0} - {1}" -f $baseTitle, $cfg.customer.name)

    [System.Windows.Forms.MessageBox]::Show("customer.json gespeichert: $path", "Gespeichert") | Out-Null
})

$btnNewCustomer.Add_Click({
    $cfg = New-DefaultCustomerConfig
    Apply-CustomerToGui -Config $cfg
    $btnNewCustomer.Visible = $false

    [System.Windows.Forms.MessageBox]::Show("Defaults geladen. Bitte speichern.", "Info") | Out-Null
})

$btnOpenCustomer.Add_Click({
    $path = Get-CustomerPath
    if (Test-Path $path) {
        Start-Process notepad.exe "`"$path`""
    } else {
        [System.Windows.Forms.MessageBox]::Show("customer.json existiert nicht.", "Hinweis") | Out-Null
    }
})

$btnOpenLogs.Add_Click({
    $relative = $txtRelativeLogPath.Text.Trim()
    if (-not $relative) { $relative = "logs" }
    $logsPath = Join-Path $toolRoot $relative
    Start-Process explorer.exe "`"$logsPath`""
})

$existingPolicy = Read-ExistingPolicy
if ($existingPolicy) { Apply-PolicyToGui $existingPolicy }

$existingCustomer = Read-ExistingCustomer
if ($existingCustomer) {
    Apply-CustomerToGui -Config $existingCustomer
    $btnNewCustomer.Visible = $false
    $form.Text = ("{0} - {1}" -f $baseTitle, $existingCustomer.customer.name)
} else {
    Apply-CustomerToGui -Config (New-DefaultCustomerConfig)
    $btnNewCustomer.Visible = $true
    $lblCfgWarning.Text = "customer.json fehlt oder ist ungueltig. Bitte Defaults laden."
    $lblCfgWarning.Visible = $true
}

[void]$form.ShowDialog()
