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

$form = New-Object System.Windows.Forms.Form
$form.Text = "Wartung Admin - Policy Builder"
$form.StartPosition = "CenterScreen"
$form.Width = 1100
$form.Height = 700
$form.MinimumSize = New-Object System.Drawing.Size(900,650)

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.ColumnCount = 1
$main.RowCount = 3
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($main)

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

$existingPolicy = Read-ExistingPolicy
if ($existingPolicy) { Apply-PolicyToGui $existingPolicy }

[void]$form.ShowDialog()
