param()

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$modulePath = Join-Path $toolRoot "shared\WartungsTools.SDK.psm1"
Import-Module $modulePath -Force

$toolManifest = Get-Content (Join-Path $toolRoot "tool.json") -Raw | ConvertFrom-Json
$toolId = $toolManifest.toolId

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Write-UiLog {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO",
        [string]$Action
    )

    Write-Log -Level $Level -Message $Message -ToolId $toolId -Trigger "AdminUI" -Action $Action
}

function Get-ActionNames {
    $actionsPath = Join-Path $toolRoot "Actions"
    if (-not (Test-Path $actionsPath)) { return @() }
    return Get-ChildItem -Path $actionsPath -Filter "*.ps1" | Sort-Object Name | ForEach-Object {
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    }
}

function New-LabeledTextBox {
    param(
        [string]$LabelText,
        [int]$TextWidth = 260,
        [switch]$Multiline
    )

    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.FlowDirection = "LeftToRight"
    $panel.WrapContents = $false
    $panel.AutoSize = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Width = 140
    $label.TextAlign = "MiddleLeft"

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Width = $TextWidth

    if ($Multiline) {
        $textbox.Multiline = $true
        $textbox.Height = 70
        $textbox.ScrollBars = "Vertical"
    }

    $panel.Controls.Add($label)
    $panel.Controls.Add($textbox)

    return [pscustomobject]@{
        Panel = $panel
        TextBox = $textbox
    }
}

function New-ActionGrid {
    param(
        [string[]]$ActionNames
    )

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 430
    $panel.Height = 180

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.AllowUserToAddRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false

    $colName = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colName.HeaderText = "Action"
    $colName.DataSource = $ActionNames
    $colName.Width = 180

    $colParams = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colParams.HeaderText = "Params (JSON)"
    $colParams.Width = 230

    [void]$grid.Columns.Add($colName)
    [void]$grid.Columns.Add($colParams)

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.FlowDirection = "LeftToRight"
    $buttons.WrapContents = $false
    $buttons.Dock = "Bottom"
    $buttons.AutoSize = $true

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add"
    $btnAdd.AutoSize = $true

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Remove"
    $btnRemove.AutoSize = $true

    $buttons.Controls.Add($btnAdd)
    $buttons.Controls.Add($btnRemove)

    $panel.Controls.Add($grid)
    $panel.Controls.Add($buttons)

    return [pscustomobject]@{
        Panel = $panel
        Grid = $grid
        AddButton = $btnAdd
        RemoveButton = $btnRemove
    }
}

function Add-ActionRow {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string[]]$ActionNames,
        [pscustomobject]$Action,
        [switch]$UseDefaultAction,
        [switch]$EmptyParamsWhenNull
    )

    $rowIndex = $Grid.Rows.Add()
    $row = $Grid.Rows[$rowIndex]

    if ($Action -and $Action.name) {
        $row.Cells[0].Value = $Action.name
    } elseif ($UseDefaultAction -and $ActionNames.Count -gt 0) {
        $row.Cells[0].Value = $ActionNames[0]
    }

    if ($Action -and $Action.params) {
        $row.Cells[1].Value = ($Action.params | ConvertTo-Json -Depth 6 -Compress)
    } elseif ($EmptyParamsWhenNull) {
        $row.Cells[1].Value = ""
    } else {
        $row.Cells[1].Value = "{}"
    }
}

function Set-ActionsToGrid {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string[]]$ActionNames,
        [object[]]$Actions
    )

    $Grid.DataSource = $null
    $Grid.Rows.Clear()
    if (-not $Actions) { return }

    foreach ($action in $Actions) {
        Add-ActionRow -Grid $Grid -ActionNames $ActionNames -Action $action -EmptyParamsWhenNull
    }
}

function Normalize-Actions {
    param([object]$Value)

    if (-not $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return @($Value)
}

function Get-TargetLines {
    param([object]$Value)

    if (-not $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return @($Value)
}

function Remove-SelectedActionRow {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Context
    )

    if ($null -eq $Grid) { return }
    $Grid.EndEdit()
    $Grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)

    $rowIndex = $null
    if ($Grid.SelectedRows.Count -gt 0) {
        $rowIndex = $Grid.SelectedRows[0].Index
    } elseif ($Grid.CurrentRow) {
        $rowIndex = $Grid.CurrentRow.Index
    } elseif ($Grid.CurrentCell) {
        $rowIndex = $Grid.CurrentCell.RowIndex
    }

    if ($null -eq $rowIndex) {
        [System.Windows.Forms.MessageBox]::Show("Bitte eine Action auswaehlen.", $Context, "OK", "Information") | Out-Null
        return
    }

    if ($rowIndex -ge 0 -and $rowIndex -lt $Grid.Rows.Count) {
        $row = $Grid.Rows[$rowIndex]
        if (-not $row.IsNewRow) {
            $Grid.Rows.RemoveAt($rowIndex)
        }
    }

    $Grid.Refresh()
}

function New-PocActionGrid {
    param(
        [string[]]$ActionNames
    )

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.AllowUserToAddRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false

    $colEnabled = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colEnabled.HeaderText = "Enabled"

    $colTrigger = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colTrigger.HeaderText = "Trigger"
    $colTrigger.DataSource = @("Logon","Logoff","Both")

    $colFrequency = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colFrequency.HeaderText = "Frequency"
    $colFrequency.DataSource = @("Every","Once")

    $colAction = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colAction.HeaderText = "Action"
    $colAction.DataSource = $ActionNames

    $colMode = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colMode.HeaderText = "Mode"
    $colMode.DataSource = @("Silent","Interactive")

    $colParams = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colParams.HeaderText = "Params (JSON)"

    $colCampaignId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colCampaignId.HeaderText = "CampaignId"

    $colValidUntil = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colValidUntil.HeaderText = "ValidUntil"

    $colTargetsUsers = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTargetsUsers.HeaderText = "TargetsUsers"

    $colTargetsGroups = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTargetsGroups.HeaderText = "TargetsGroups"

    [void]$grid.Columns.Add($colEnabled)
    [void]$grid.Columns.Add($colTrigger)
    [void]$grid.Columns.Add($colFrequency)
    [void]$grid.Columns.Add($colAction)
    [void]$grid.Columns.Add($colMode)
    [void]$grid.Columns.Add($colParams)
    [void]$grid.Columns.Add($colCampaignId)
    [void]$grid.Columns.Add($colValidUntil)
    [void]$grid.Columns.Add($colTargetsUsers)
    [void]$grid.Columns.Add($colTargetsGroups)

    return $grid
}

function Get-ActionsFromGrid {
    param(
        [System.Windows.Forms.DataGridView]$Grid
    )

    $actions = @()
    foreach ($row in $Grid.Rows) {
        $name = [string]$row.Cells[0].Value
        if (-not $name) { continue }

        $paramsText = [string]$row.Cells[1].Value
        $params = @{}
        if ($paramsText) {
            try {
                $parsed = $paramsText | ConvertFrom-Json
                if ($null -eq $parsed) {
                    $params = @{}
                } elseif ($parsed -is [System.Collections.IDictionary] -or $parsed -is [pscustomobject]) {
                    $params = $parsed
                } else {
                    throw "Params must be a JSON object."
                }
            } catch {
                throw "Invalid params JSON for action '$name': $($_.Exception.Message)"
            }
        }

        $actions += [pscustomobject]@{
            name = $name
            mode = "Silent"
            params = $params
        }
    }

    return $actions
}

function Get-Lines {
    param([string]$Text)

    if (-not $Text) { return @() }
    return $Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Normalize-Once {
    param([object]$Value)

    if (-not $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return @($Value)
}

function Normalize-Every {
    param([object]$Value)

    if (-not $Value) { return $null }
    if ($Value -is [System.Array]) {
        if ($Value.Count -gt 0) { return $Value[0] }
        return $null
    }
    return $Value
}

function Get-DefaultPolicy {
    return [pscustomobject]@{
        logon = [pscustomobject]@{
            every = [pscustomobject]@{ enabled = $false; actions = @() }
            once = @()
        }
        logoff = [pscustomobject]@{
            every = [pscustomobject]@{ enabled = $false; actions = @() }
            once = @()
        }
    }
}

function Get-DefaultCustomer {
    return [pscustomobject]@{
        customer = [pscustomobject]@{ name = "" }
        paths = [pscustomobject]@{ repoRoot = "" }
        fslogix = [pscustomobject]@{ enabled = $false; profileShare = ""; officeContainerShare = "" }
        branding = [pscustomobject]@{ windowTitle = "Wartung Admin"; supportText = "" }
        logging = [pscustomobject]@{ relativeLogPath = "logs"; adminLogRoot = "" }
        flags = [pscustomobject]@{ allowOffline = $true; allowLogoffRunner = $true }
    }
}

$autoLoadPolicy = $true

function Get-ToolRootPath {
    return $toolRoot
}

function Get-RecentCampaignIdsPath {
    param(
        [string]$ToolRoot
    )
    return (Join-Path $ToolRoot "recent_campaignIds.json")
}

function Load-RecentCampaignIds {
    param(
        [string]$ToolRoot
    )
    $path = Get-RecentCampaignIdsPath -ToolRoot $ToolRoot
    Write-UiLog ("RecentCampaignIds path: {0}" -f $path) "INFO"
    if (-not (Test-Path $path)) { return @() }
    try {
        $items = Get-Content $path -Raw | ConvertFrom-Json
        $list = @()
        if ($items -is [string]) {
            $value = $items.Trim()
            if ($value) { $list += $value }
            return $list
        }
        if ($items -is [System.Collections.IEnumerable]) {
            foreach ($item in $items) {
                if ($null -eq $item) { continue }
                $value = [string]$item
                if ($value.Trim()) { $list += $value.Trim() }
            }
            return $list
        }
        return @()
    } catch {
        Write-UiLog ("RecentCampaignIds: invalid JSON, ignoring file: {0}" -f $path) "WARN"
        return @()
    }
}

function Save-RecentCampaignIds {
    param(
        [string]$ToolRoot,
        [string]$CampaignId
    )

    $path = Get-RecentCampaignIdsPath -ToolRoot $ToolRoot
    $newId = if ($CampaignId) { $CampaignId.Trim() } else { "" }
    if (-not $newId) { return }

    $existing = Load-RecentCampaignIds -ToolRoot $ToolRoot
    $recentList = @()
    $recentMap = @{}
    $recentList += $newId
    $recentMap[$newId.ToLowerInvariant()] = $true

    foreach ($item in $existing) {
        $value = [string]$item
        if (-not $value) { continue }
        $trimmed = $value.Trim()
        if (-not $trimmed) { continue }
        $key = $trimmed.ToLowerInvariant()
        if ($recentMap.ContainsKey($key)) { continue }
        $recentList += $trimmed
        $recentMap[$key] = $true
    }

    if ($recentList.Count -gt 10) {
        $recentList = $recentList[0..9]
    }

    $json = @($recentList) | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Wartung Admin"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(980, 760)
$form.MaximizeBox = $true

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"

$tabPolicy = New-Object System.Windows.Forms.TabPage
$tabPolicy.Text = "Policy"

$tabConfig = New-Object System.Windows.Forms.TabPage
$tabConfig.Text = "Konfiguration"

$tabs.Controls.Add($tabPolicy)
$tabs.Controls.Add($tabConfig)

$actionNames = Get-ActionNames

$policyRoot = New-Object System.Windows.Forms.TableLayoutPanel
$policyRoot.Dock = "Fill"
$policyRoot.RowCount = 2
$policyRoot.ColumnCount = 1
$policyRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$policyRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$grpLogonOnce = New-Object System.Windows.Forms.GroupBox
$grpLogonOnce.Text = "Logon Once"
$grpLogonOnce.Dock = "Fill"

$grpLogonEvery = New-Object System.Windows.Forms.GroupBox
$grpLogonEvery.Text = "Logon Every"
$grpLogonEvery.Dock = "Fill"

$grpLogoffOnce = New-Object System.Windows.Forms.GroupBox
$grpLogoffOnce.Text = "Logoff Once"
$grpLogoffOnce.Dock = "Fill"

$grpLogoffEvery = New-Object System.Windows.Forms.GroupBox
$grpLogoffEvery.Text = "Logoff Every"
$grpLogoffEvery.Dock = "Fill"

$policyTabs = New-Object System.Windows.Forms.TabControl
$policyTabs.Dock = "Fill"

$tabLogon = New-Object System.Windows.Forms.TabPage
$tabLogon.Text = "Logon"

$tabLogoff = New-Object System.Windows.Forms.TabPage
$tabLogoff.Text = "Logoff"

$tabActionsPoc = New-Object System.Windows.Forms.TabPage
$tabActionsPoc.Text = "Actions (PoC)"

$tlpLogon = New-Object System.Windows.Forms.TableLayoutPanel
$tlpLogon.Dock = "Fill"
$tlpLogon.ColumnCount = 1
$tlpLogon.RowCount = 2
$tlpLogon.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tlpLogon.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$tlpLogoff = New-Object System.Windows.Forms.TableLayoutPanel
$tlpLogoff.Dock = "Fill"
$tlpLogoff.ColumnCount = 1
$tlpLogoff.RowCount = 2
$tlpLogoff.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tlpLogoff.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$tlpLogon.Controls.Add($grpLogonEvery, 0, 0)
$tlpLogon.Controls.Add($grpLogonOnce, 0, 1)
$tlpLogoff.Controls.Add($grpLogoffEvery, 0, 0)
$tlpLogoff.Controls.Add($grpLogoffOnce, 0, 1)

$tabLogon.Controls.Add($tlpLogon)
$tabLogoff.Controls.Add($tlpLogoff)

$policyTabs.Controls.Add($tabLogon)
$policyTabs.Controls.Add($tabLogoff)
$policyTabs.Controls.Add($tabActionsPoc)

$policyRoot.Controls.Add($policyTabs, 0, 0)

$policyButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$policyButtons.FlowDirection = "LeftToRight"
$policyButtons.WrapContents = $false
$policyButtons.AutoSize = $true
$policyButtons.Dock = "Fill"

$btnPolicyLoad = New-Object System.Windows.Forms.Button
$btnPolicyLoad.Text = "Vorhandene Policy laden"
$btnPolicyLoad.AutoSize = $true

$btnPolicySave = New-Object System.Windows.Forms.Button
$btnPolicySave.Text = "Policy schreiben"
$btnPolicySave.AutoSize = $true

$lblPolicyStatus = New-Object System.Windows.Forms.Label
$lblPolicyStatus.AutoSize = $true
$lblPolicyStatus.Text = ""
$lblPolicyStatus.Padding = "10,8,0,0"

$policyButtons.Controls.Add($btnPolicyLoad)
$policyButtons.Controls.Add($btnPolicySave)
$policyButtons.Controls.Add($lblPolicyStatus)

$policyRoot.Controls.Add($policyButtons, 0, 1)
$tabPolicy.Controls.Add($policyRoot)

$logonOncePanel = New-Object System.Windows.Forms.FlowLayoutPanel
$logonOncePanel.Dock = "Fill"
$logonOncePanel.FlowDirection = "TopDown"
$logonOncePanel.WrapContents = $false
$logonOncePanel.AutoScroll = $true

$chkLogonOnceEnabled = New-Object System.Windows.Forms.CheckBox
$chkLogonOnceEnabled.Text = "Enabled"
$chkLogonOnceEnabled.Checked = $false
$chkLogonOnceEnabled.AutoCheck = $true
$chkLogonOnceEnabled.Enabled = $true

$logonOnceCampaign = New-LabeledTextBox -LabelText "CampaignId" -TextWidth 240
$logonOnceValidUntil = New-LabeledTextBox -LabelText "ValidUntil (YYYY-MM-DD)" -TextWidth 140
$logonOnceUsers = New-LabeledTextBox -LabelText "Targets Users" -TextWidth 240 -Multiline
$logonOnceGroups = New-LabeledTextBox -LabelText "Targets Groups" -TextWidth 240 -Multiline
$logonOnceActions = New-ActionGrid -ActionNames $actionNames

$logonOncePanel.Controls.Add($chkLogonOnceEnabled)
$logonOncePanel.Controls.Add($logonOnceCampaign.Panel)
$logonOncePanel.Controls.Add($logonOnceValidUntil.Panel)
$logonOncePanel.Controls.Add($logonOnceUsers.Panel)
$logonOncePanel.Controls.Add($logonOnceGroups.Panel)
$logonOncePanel.Controls.Add($logonOnceActions.Panel)

$grpLogonOnce.Controls.Add($logonOncePanel)

$logonEveryPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$logonEveryPanel.Dock = "Fill"
$logonEveryPanel.FlowDirection = "TopDown"
$logonEveryPanel.WrapContents = $false
$logonEveryPanel.AutoScroll = $true

$chkLogonEveryEnabled = New-Object System.Windows.Forms.CheckBox
$chkLogonEveryEnabled.Text = "Enabled"

$logonEveryActions = New-ActionGrid -ActionNames $actionNames

$logonEveryPanel.Controls.Add($chkLogonEveryEnabled)
$logonEveryPanel.Controls.Add($logonEveryActions.Panel)

$grpLogonEvery.Controls.Add($logonEveryPanel)

$logoffOncePanel = New-Object System.Windows.Forms.FlowLayoutPanel
$logoffOncePanel.Dock = "Fill"
$logoffOncePanel.FlowDirection = "TopDown"
$logoffOncePanel.WrapContents = $false
$logoffOncePanel.AutoScroll = $true

$chkLogoffOnceEnabled = New-Object System.Windows.Forms.CheckBox
$chkLogoffOnceEnabled.Text = "Enabled"
$chkLogoffOnceEnabled.Checked = $false
$chkLogoffOnceEnabled.AutoCheck = $true
$chkLogoffOnceEnabled.Enabled = $true

$logoffOnceCampaign = New-LabeledTextBox -LabelText "CampaignId" -TextWidth 240
$logoffOnceValidUntil = New-LabeledTextBox -LabelText "ValidUntil (YYYY-MM-DD)" -TextWidth 140
$logoffOnceUsers = New-LabeledTextBox -LabelText "Targets Users" -TextWidth 240 -Multiline
$logoffOnceGroups = New-LabeledTextBox -LabelText "Targets Groups" -TextWidth 240 -Multiline
$logoffOnceActions = New-ActionGrid -ActionNames $actionNames

$logoffOncePanel.Controls.Add($chkLogoffOnceEnabled)
$logoffOncePanel.Controls.Add($logoffOnceCampaign.Panel)
$logoffOncePanel.Controls.Add($logoffOnceValidUntil.Panel)
$logoffOncePanel.Controls.Add($logoffOnceUsers.Panel)
$logoffOncePanel.Controls.Add($logoffOnceGroups.Panel)
$logoffOncePanel.Controls.Add($logoffOnceActions.Panel)

$grpLogoffOnce.Controls.Add($logoffOncePanel)

$logoffEveryPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$logoffEveryPanel.Dock = "Fill"
$logoffEveryPanel.FlowDirection = "TopDown"
$logoffEveryPanel.WrapContents = $false
$logoffEveryPanel.AutoScroll = $true

$chkLogoffEveryEnabled = New-Object System.Windows.Forms.CheckBox
$chkLogoffEveryEnabled.Text = "Enabled"

$logoffEveryActions = New-ActionGrid -ActionNames $actionNames

$logoffEveryPanel.Controls.Add($chkLogoffEveryEnabled)
$logoffEveryPanel.Controls.Add($logoffEveryActions.Panel)

$grpLogoffEvery.Controls.Add($logoffEveryPanel)

$pocPanel = New-Object System.Windows.Forms.Panel
$pocPanel.Dock = "Fill"

$chkPocShowEnabled = New-Object System.Windows.Forms.CheckBox
$chkPocShowEnabled.Text = "Nur Enabled anzeigen"
$chkPocShowEnabled.Checked = $false
$chkPocShowEnabled.AutoSize = $true
$chkPocShowEnabled.Dock = "Top"

$gridActionsPoc = New-PocActionGrid -ActionNames $actionNames
$gridActionsPoc.Dock = "Fill"

$pocButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pocButtons.FlowDirection = "LeftToRight"
$pocButtons.WrapContents = $false
$pocButtons.AutoSize = $true
$pocButtons.Dock = "Bottom"

$btnPocAdd = New-Object System.Windows.Forms.Button
$btnPocAdd.Text = "Add"
$btnPocAdd.AutoSize = $true

$btnPocRemove = New-Object System.Windows.Forms.Button
$btnPocRemove.Text = "Remove"
$btnPocRemove.AutoSize = $true

$btnPocDuplicate = New-Object System.Windows.Forms.Button
$btnPocDuplicate.Text = "Duplicate"
$btnPocDuplicate.AutoSize = $true

$btnPocDisableCampaign = New-Object System.Windows.Forms.Button
$btnPocDisableCampaign.Text = "Disable Campaign"
$btnPocDisableCampaign.AutoSize = $true

$btnPocQuickAddOnce = New-Object System.Windows.Forms.Button
$btnPocQuickAddOnce.Text = "Quick Add Once Campaign..."
$btnPocQuickAddOnce.AutoSize = $true

$btnPocPreview = New-Object System.Windows.Forms.Button
$btnPocPreview.Text = "Preview (Was wuerde laufen?)"
$btnPocPreview.AutoSize = $true

$btnPocParamsCheck = New-Object System.Windows.Forms.Button
$btnPocParamsCheck.Text = "Params pruefen"
$btnPocParamsCheck.AutoSize = $true

$btnPocParamsFormat = New-Object System.Windows.Forms.Button
$btnPocParamsFormat.Text = "Params formatieren"
$btnPocParamsFormat.AutoSize = $true

$chkPolicyUsePoc = New-Object System.Windows.Forms.CheckBox
$chkPolicyUsePoc.Text = "PoC als Quelle verwenden"
$chkPolicyUsePoc.Checked = $true
$chkPolicyUsePoc.AutoSize = $true

$pocButtons.Controls.Add($btnPocAdd)
$pocButtons.Controls.Add($btnPocRemove)
$pocButtons.Controls.Add($btnPocDuplicate)
$pocButtons.Controls.Add($btnPocDisableCampaign)
$pocButtons.Controls.Add($btnPocQuickAddOnce)
$pocButtons.Controls.Add($btnPocPreview)
$pocButtons.Controls.Add($btnPocParamsCheck)
$pocButtons.Controls.Add($btnPocParamsFormat)

$pocPanel.Controls.Add($gridActionsPoc)
$pocPanel.Controls.Add($chkPocShowEnabled)
$pocPanel.Controls.Add($pocButtons)

$tabActionsPoc.Controls.Add($pocPanel)

$logonEveryActions.AddButton.Add_Click({ Add-ActionRow -Grid $logonEveryActions.Grid -ActionNames $actionNames -UseDefaultAction })
$logonEveryActions.RemoveButton.Add_Click({
    Remove-SelectedActionRow -Grid $logonEveryActions.Grid -Context "Logon Every"
})

$logonOnceActions.AddButton.Add_Click({ Add-ActionRow -Grid $logonOnceActions.Grid -ActionNames $actionNames -UseDefaultAction })
$logonOnceActions.RemoveButton.Add_Click({
    Remove-SelectedActionRow -Grid $logonOnceActions.Grid -Context "Logon Once"
})

$logoffEveryActions.AddButton.Add_Click({ Add-ActionRow -Grid $logoffEveryActions.Grid -ActionNames $actionNames -UseDefaultAction })
$logoffEveryActions.RemoveButton.Add_Click({
    Remove-SelectedActionRow -Grid $logoffEveryActions.Grid -Context "Logoff Every"
})

$logoffOnceActions.AddButton.Add_Click({ Add-ActionRow -Grid $logoffOnceActions.Grid -ActionNames $actionNames -UseDefaultAction })
$logoffOnceActions.RemoveButton.Add_Click({
    Remove-SelectedActionRow -Grid $logoffOnceActions.Grid -Context "Logoff Once"
})

$btnPocAdd.Add_Click({
    Add-PocRow -Grid $gridActionsPoc -RowData @{
        Enabled = $true
        Trigger = "Logon"
        Frequency = "Every"
        Action = if ($actionNames.Count -gt 0) { $actionNames[0] } else { "" }
        Mode = "Silent"
        Params = "{}"
        CampaignId = ""
        ValidUntil = ""
        TargetsUsers = ""
        TargetsGroups = ""
    }
    $lastIndex = $gridActionsPoc.Rows.Count - 1
    if ($lastIndex -ge 0) {
        $gridActionsPoc.CurrentCell = $gridActionsPoc.Rows[$lastIndex].Cells[3]
        $gridActionsPoc.BeginEdit($true)
    }
})

$btnPocRemove.Add_Click({
    Remove-SelectedActionRow -Grid $gridActionsPoc -Context "Actions (PoC)"
})

$btnPocDuplicate.Add_Click({
    if ($gridActionsPoc.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte eine Zeile auswählen.", "Hinweis", "OK", "Information") | Out-Null
        return
    }

    $selectedRow = $gridActionsPoc.SelectedRows[0]
    if ($selectedRow.IsNewRow) { return }

    $insertIndex = $selectedRow.Index + 1
    $newIndex = $gridActionsPoc.Rows.Insert($insertIndex, 1)
    $newRow = $gridActionsPoc.Rows[$insertIndex]

    for ($col = 0; $col -lt $gridActionsPoc.Columns.Count; $col += 1) {
        $newRow.Cells[$col].Value = $selectedRow.Cells[$col].Value
    }

    $gridActionsPoc.CurrentCell = $newRow.Cells[3]
    $gridActionsPoc.BeginEdit($true)
})

$btnPocDisableCampaign.Add_Click({
    if ($gridActionsPoc.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte eine Zeile auswählen.", "Hinweis", "OK", "Information") | Out-Null
        return
    }

    $selectedRow = $gridActionsPoc.SelectedRows[0]
    if ($selectedRow.IsNewRow) { return }

    $campaignId = [string]$selectedRow.Cells[6].Value
    if ([string]::IsNullOrWhiteSpace($campaignId)) {
        [System.Windows.Forms.MessageBox]::Show("CampaignId ist leer.", "Hinweis", "OK", "Information") | Out-Null
        return
    }

    $campaignKey = $campaignId.Trim().ToLowerInvariant()
    foreach ($row in $gridActionsPoc.Rows) {
        if ($row.IsNewRow) { continue }
        $rowCampaign = [string]$row.Cells[6].Value
        if ($rowCampaign -and $rowCampaign.Trim().ToLowerInvariant() -eq $campaignKey) {
            $row.Cells[0].Value = $false
        }
    }

    $chkPocShowEnabled.Checked = $chkPocShowEnabled.Checked
})

$btnPocQuickAddOnce.Add_Click({
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Quick Add Once Campaign"
    $dialog.StartPosition = "CenterParent"
    $dialog.Size = New-Object System.Drawing.Size(760, 600)
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $dlgPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $dlgPanel.Dock = "Fill"
    $dlgPanel.ColumnCount = 2
    $dlgPanel.RowCount = 7
    $dlgPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 35)))
    $dlgPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 65)))
    for ($i = 0; $i -lt 7; $i += 1) {
        $dlgPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    }

    $lblCampaign = New-Object System.Windows.Forms.Label
    $lblCampaign.Text = "CampaignId"
    $lblCampaign.AutoSize = $true
    $txtCampaign = New-Object System.Windows.Forms.ComboBox
    $txtCampaign.Width = 420
    $txtCampaign.DropDownStyle = "DropDown"
    $txtCampaign.AutoCompleteMode = "SuggestAppend"
    $txtCampaign.AutoCompleteSource = "ListItems"
    $recentCampaignIds = Load-RecentCampaignIds -ToolRoot (Get-ToolRootPath)
    $txtCampaign.Items.Clear()
    foreach ($item in $recentCampaignIds) { [void]$txtCampaign.Items.Add($item) }

    $lblTrigger = New-Object System.Windows.Forms.Label
    $lblTrigger.Text = "Trigger"
    $lblTrigger.AutoSize = $true
    $cmbTrigger = New-Object System.Windows.Forms.ComboBox
    $cmbTrigger.DropDownStyle = "DropDownList"
    [void]$cmbTrigger.Items.AddRange(@("Logon","Logoff","Both"))
    $cmbTrigger.SelectedItem = "Logon"

    $lblValidUntil = New-Object System.Windows.Forms.Label
    $lblValidUntil.Text = "ValidUntil (yyyy-MM-dd)"
    $lblValidUntil.AutoSize = $true
    $txtValidUntil = New-Object System.Windows.Forms.TextBox
    $txtValidUntil.Width = 240

    $lblUsers = New-Object System.Windows.Forms.Label
    $lblUsers.Text = "Targets Users"
    $lblUsers.AutoSize = $true
    $txtUsers = New-Object System.Windows.Forms.TextBox
    $txtUsers.Multiline = $true
    $txtUsers.Height = 70
    $txtUsers.ScrollBars = "Vertical"

    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = "Targets Groups"
    $lblGroups.AutoSize = $true
    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Multiline = $true
    $txtGroups.Height = 70
    $txtGroups.ScrollBars = "Vertical"

    $lblActions = New-Object System.Windows.Forms.Label
    $lblActions.Text = "Actions"
    $lblActions.AutoSize = $true

    $lblActionFilter = New-Object System.Windows.Forms.Label
    $lblActionFilter.Text = "Filter Actions:"
    $lblActionFilter.AutoSize = $true

    $txtActionFilter = New-Object System.Windows.Forms.TextBox
    $txtActionFilter.Width = 520

    $lstActions = New-Object System.Windows.Forms.CheckedListBox
    $lstActions.CheckOnClick = $true
    $lstActions.Height = 170
    $lstActions.Anchor = "Left,Right,Top"

    $allActions = @($actionNames)
    $checkedMap = @{}

    function Refresh-ActionFilterList {
        $filter = $txtActionFilter.Text.Trim().ToLowerInvariant()
        $lstActions.BeginUpdate()
        try {
            $lstActions.Items.Clear()
            foreach ($name in $allActions) {
                if ($filter -and ($name.ToLowerInvariant() -notlike "*$filter*")) { continue }
                $index = $lstActions.Items.Add($name)
                if ($checkedMap.ContainsKey($name) -and $checkedMap[$name]) {
                    $lstActions.SetItemChecked($index, $true)
                }
            }
        } finally {
            $lstActions.EndUpdate()
        }
    }

    $lstActions.Add_ItemCheck({
        $name = [string]$lstActions.Items[$_.Index]
        if ($name) {
            $checkedMap[$name] = ($_.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
        }
    })

    $txtActionFilter.Add_TextChanged({
        Refresh-ActionFilterList
    })

    $txtActionFilter.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $txtActionFilter.Text = ""
            $_.Handled = $true
        }
    })

    Refresh-ActionFilterList

    $lblMode = New-Object System.Windows.Forms.Label
    $lblMode.Text = "Mode"
    $lblMode.AutoSize = $true
    $cmbMode = New-Object System.Windows.Forms.ComboBox
    $cmbMode.DropDownStyle = "DropDownList"
    [void]$cmbMode.Items.AddRange(@("Silent","Interactive"))
    $cmbMode.SelectedItem = "Silent"

    $dlgButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $dlgButtons.FlowDirection = "RightToLeft"
    $dlgButtons.WrapContents = $false
    $dlgButtons.AutoSize = $true

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.AutoSize = $true

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.AutoSize = $true

    $dlgButtons.Controls.Add($btnOk)
    $dlgButtons.Controls.Add($btnCancel)

    $dlgPanel.Controls.Add($lblCampaign, 0, 0)
    $dlgPanel.Controls.Add($txtCampaign, 1, 0)
    $dlgPanel.Controls.Add($lblTrigger, 0, 1)
    $dlgPanel.Controls.Add($cmbTrigger, 1, 1)
    $dlgPanel.Controls.Add($lblValidUntil, 0, 2)
    $dlgPanel.Controls.Add($txtValidUntil, 1, 2)
    $dlgPanel.Controls.Add($lblUsers, 0, 3)
    $dlgPanel.Controls.Add($txtUsers, 1, 3)
    $dlgPanel.Controls.Add($lblGroups, 0, 4)
    $dlgPanel.Controls.Add($txtGroups, 1, 4)
    $actionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $actionsPanel.FlowDirection = "TopDown"
    $actionsPanel.WrapContents = $false
    $actionsPanel.AutoSize = $true
    $actionsPanel.Dock = "Fill"
    $actionsPanel.Controls.Add($lblActionFilter)
    $actionsPanel.Controls.Add($txtActionFilter)
    $actionsPanel.Controls.Add($lstActions)

    $dlgPanel.Controls.Add($lblActions, 0, 5)
    $dlgPanel.Controls.Add($actionsPanel, 1, 5)
    $dlgPanel.Controls.Add($lblMode, 0, 6)
    $dlgPanel.Controls.Add($cmbMode, 1, 6)

    $dialog.Controls.Add($dlgPanel)
    $dialog.Controls.Add($dlgButtons)
    $dlgButtons.Dock = "Bottom"

    $btnCancel.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $btnOk.Add_Click({
        $campaignId = $txtCampaign.Text.Trim()
        if (-not $campaignId) {
            [System.Windows.Forms.MessageBox]::Show("CampaignId ist erforderlich.", "Validierung", "OK", "Warning") | Out-Null
            return
        }
        if ($lstActions.CheckedItems.Count -lt 1) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Action auswählen.", "Validierung", "OK", "Warning") | Out-Null
            return
        }

        $trigger = [string]$cmbTrigger.SelectedItem
        $mode = [string]$cmbMode.SelectedItem
        $validUntil = $txtValidUntil.Text.Trim()
        $targetsUsers = $txtUsers.Text
        $targetsGroups = $txtGroups.Text

        $insertedIndices = @()

        foreach ($actionName in $lstActions.CheckedItems) {
            $triggers = if ($trigger -eq "Both") { @("Logon","Logoff") } else { @($trigger) }
            foreach ($t in $triggers) {
                Add-PocRow -Grid $gridActionsPoc -RowData @{
                    Enabled = $true
                    Trigger = $t
                    Frequency = "Once"
                    Action = [string]$actionName
                    Mode = $mode
                    Params = "{}"
                    CampaignId = $campaignId
                    ValidUntil = $validUntil
                    TargetsUsers = $targetsUsers
                    TargetsGroups = $targetsGroups
                }
                $insertedIndices += ($gridActionsPoc.Rows.Count - 1)
            }
        }

        Save-RecentCampaignIds -ToolRoot (Get-ToolRootPath) -CampaignId $campaignId

        if ($chkPocShowEnabled.Checked) {
            $chkPocShowEnabled.Checked = $chkPocShowEnabled.Checked
        }

        if ($insertedIndices.Count -gt 0) {
            $firstIndex = $insertedIndices[0]
            $gridActionsPoc.ClearSelection()
            $gridActionsPoc.Rows[$firstIndex].Selected = $true
            $gridActionsPoc.CurrentCell = $gridActionsPoc.Rows[$firstIndex].Cells[3]
        }

        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    [void]$dialog.ShowDialog($form)
})

function Get-PocCurrentRow {
    if ($gridActionsPoc.SelectedRows.Count -gt 0) {
        return $gridActionsPoc.SelectedRows[0]
    }
    if ($gridActionsPoc.CurrentRow) { return $gridActionsPoc.CurrentRow }
    if ($gridActionsPoc.CurrentCell) { return $gridActionsPoc.Rows[$gridActionsPoc.CurrentCell.RowIndex] }
    return $null
}

function Set-PocParamsCellStyle {
    param(
        [System.Windows.Forms.DataGridViewCell]$Cell,
        [switch]$IsValid
    )

    if ($null -eq $Cell) { return }
    if ($IsValid) {
        $Cell.Style.BackColor = $gridActionsPoc.DefaultCellStyle.BackColor
    } else {
        $Cell.Style.BackColor = [System.Drawing.Color]::MistyRose
    }
}

function Test-PocParamsJson {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return @($true, @{}) }
    try {
        $parsed = $Text | ConvertFrom-Json -ErrorAction Stop
        return @($true, $parsed)
    } catch {
        return @($false, $_.Exception.Message)
    }
}

$btnPocParamsCheck.Add_Click({
    $row = Get-PocCurrentRow
    if ($null -eq $row -or $row.IsNewRow) {
        [System.Windows.Forms.MessageBox]::Show("Bitte eine Zeile auswählen.", "Hinweis", "OK", "Information") | Out-Null
        return
    }

    $cell = $row.Cells[5]
    $text = [string]$cell.Value
    $result = Test-PocParamsJson -Text $text
    $isValid = [bool]$result[0]
    if ($isValid) {
        Set-PocParamsCellStyle -Cell $cell -IsValid
        [System.Windows.Forms.MessageBox]::Show("Params: OK", "Validierung", "OK", "Information") | Out-Null
        return
    }

    Set-PocParamsCellStyle -Cell $cell
    [System.Windows.Forms.MessageBox]::Show("Params: ungültiges JSON. " + [string]$result[1], "Validierung", "OK", "Warning") | Out-Null
})

$btnPocParamsFormat.Add_Click({
    $row = Get-PocCurrentRow
    if ($null -eq $row -or $row.IsNewRow) {
        [System.Windows.Forms.MessageBox]::Show("Bitte eine Zeile auswählen.", "Hinweis", "OK", "Information") | Out-Null
        return
    }

    $cell = $row.Cells[5]
    $text = [string]$cell.Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        $cell.Value = "{}"
        Set-PocParamsCellStyle -Cell $cell -IsValid
        return
    }

    $result = Test-PocParamsJson -Text $text
    $isValid = [bool]$result[0]
    if (-not $isValid) {
        Set-PocParamsCellStyle -Cell $cell
        [System.Windows.Forms.MessageBox]::Show("Params: ungültiges JSON. " + [string]$result[1], "Validierung", "OK", "Warning") | Out-Null
        return
    }

    $formatted = $result[1] | ConvertTo-Json -Depth 10
    $cell.Value = $formatted
    Set-PocParamsCellStyle -Cell $cell -IsValid
})

$gridActionsPoc.Add_CellEndEdit({
    if ($_.ColumnIndex -ne 5) { return }
    $cell = $gridActionsPoc.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
    $text = [string]$cell.Value
    $result = Test-PocParamsJson -Text $text
    $isValid = [bool]$result[0]
    if ($isValid) {
        Set-PocParamsCellStyle -Cell $cell -IsValid
    } else {
        Set-PocParamsCellStyle -Cell $cell
    }
})

function Get-PreviewPolicyFromGrid {
    param(
        [System.Windows.Forms.DataGridViewRowCollection]$Rows,
        [switch]$IncludeDisabled
    )

    $policy = Get-DefaultPolicy
    $logonEveryActions = @()
    $logoffEveryActions = @()
    $logonOnceGroups = @{}
    $logoffOnceGroups = @{}

    foreach ($row in $Rows) {
        if ($row.IsNewRow) { continue }
        $enabled = [bool]$row.Cells[0].Value
        if (-not $enabled -and -not $IncludeDisabled) { continue }

        $trigger = [string]$row.Cells[1].Value
        $frequency = [string]$row.Cells[2].Value
        $actionName = [string]$row.Cells[3].Value
        $mode = [string]$row.Cells[4].Value
        $paramsText = [string]$row.Cells[5].Value
        $campaignId = [string]$row.Cells[6].Value
        $validUntil = [string]$row.Cells[7].Value
        $targetsUsersText = [string]$row.Cells[8].Value
        $targetsGroupsText = [string]$row.Cells[9].Value

        if (-not $actionName) { continue }
        if (-not $mode) { $mode = "Silent" }

        $params = @{}
        if ($paramsText) {
            try {
                $parsed = $paramsText | ConvertFrom-Json -ErrorAction Stop
                if ($parsed -is [System.Collections.IDictionary] -or $parsed -is [pscustomobject]) {
                    $params = $parsed
                }
            } catch {
                $params = @{}
            }
        }

        $actionObj = [pscustomobject]@{
            name = $actionName
            mode = $mode
            params = $params
            enabled = $enabled
        }

        if ($frequency -eq "Every") {
            if ($trigger -eq "Logon" -or $trigger -eq "Both") {
                $logonEveryActions += $actionObj
            }
            if ($trigger -eq "Logoff" -or $trigger -eq "Both") {
                $logoffEveryActions += $actionObj
            }
            continue
        }

        if ($frequency -ne "Once") { continue }

        $users = Split-TargetText -Text $targetsUsersText
        $groups = Split-TargetText -Text $targetsGroupsText
        if ($null -eq $users) { $users = @() }
        if ($null -eq $groups) { $groups = @() }
        $key = ($campaignId + "|" + $validUntil + "|" + ($users -join ";") + "|" + ($groups -join ";"))

        if ($trigger -eq "Logon" -or $trigger -eq "Both") {
            if (-not $logonOnceGroups.ContainsKey($key)) {
                $logonOnceGroups[$key] = [pscustomobject]@{
                    enabled = $enabled
                    campaignId = $campaignId
                    validUntil = $validUntil
                    targets = [pscustomobject]@{ users = $users; groups = $groups }
                    actions = @()
                }
            }
            $logonOnceGroups[$key].actions += $actionObj
        }

        if ($trigger -eq "Logoff" -or $trigger -eq "Both") {
            if (-not $logoffOnceGroups.ContainsKey($key)) {
                $logoffOnceGroups[$key] = [pscustomobject]@{
                    enabled = $enabled
                    campaignId = $campaignId
                    validUntil = $validUntil
                    targets = [pscustomobject]@{ users = $users; groups = $groups }
                    actions = @()
                }
            }
            $logoffOnceGroups[$key].actions += $actionObj
        }
    }

    $policy.logon.every.actions = $logonEveryActions
    $policy.logoff.every.actions = $logoffEveryActions
    $policy.logon.every.enabled = ($logonEveryActions.Count -gt 0)
    $policy.logoff.every.enabled = ($logoffEveryActions.Count -gt 0)
    $policy.logon.once = @($logonOnceGroups.Values)
    $policy.logoff.once = @($logoffOnceGroups.Values)

    return $policy
}

function Test-UserTargetMatch {
    param(
        [string]$UserInput,
        [object]$Targets
    )

    if (-not $Targets) { return $true }
    $users = @($Targets.users)
    if ($users.Count -eq 0) { return $true }

    $userLower = $UserInput.ToLowerInvariant()
    $userNameOnly = $UserInput
    if ($UserInput -like "*\\*") {
        $parts = $UserInput.Split("\\", 2)
        if ($parts.Count -eq 2) { $userNameOnly = $parts[1] }
    }
    $userNameOnly = $userNameOnly.ToLowerInvariant()

    foreach ($entry in $users) {
        $value = [string]$entry
        if (-not $value) { continue }
        $valueLower = $value.ToLowerInvariant()
        if ($valueLower -eq $userLower -or $valueLower -eq $userNameOnly) {
            return $true
        }
    }

    return $false
}

$btnPocPreview.Add_Click({
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Preview"
    $dialog.StartPosition = "CenterParent"
    $dialog.Size = New-Object System.Drawing.Size(800, 550)
    $dialog.FormBorderStyle = "Sizable"
    $dialog.MaximizeBox = $true
    $dialog.MinimizeBox = $false

    $tlp = New-Object System.Windows.Forms.TableLayoutPanel
    $tlp.Dock = "Fill"
    $tlp.ColumnCount = 2
    $tlp.RowCount = 3
    $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $lblTrigger = New-Object System.Windows.Forms.Label
    $lblTrigger.Text = "Trigger"
    $lblTrigger.AutoSize = $true
    $cmbTrigger = New-Object System.Windows.Forms.ComboBox
    $cmbTrigger.DropDownStyle = "DropDownList"
    [void]$cmbTrigger.Items.AddRange(@("Logon","Logoff"))
    $cmbTrigger.SelectedItem = "Logon"

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = "User"
    $lblUser.AutoSize = $true
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Width = 320
    $txtUser.Text = if ($env:USERDOMAIN) { "$env:USERDOMAIN\\$env:USERNAME" } else { $env:USERNAME }

    $chkIncludeDisabled = New-Object System.Windows.Forms.CheckBox
    $chkIncludeDisabled.Text = "Include disabled rows"
    $chkIncludeDisabled.Checked = $false
    $chkIncludeDisabled.AutoSize = $true

    $inputsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $inputsPanel.FlowDirection = "LeftToRight"
    $inputsPanel.WrapContents = $false
    $inputsPanel.AutoSize = $true
    $inputsPanel.Dock = "Fill"
    $inputsPanel.Controls.Add($lblTrigger)
    $inputsPanel.Controls.Add($cmbTrigger)
    $inputsPanel.Controls.Add($lblUser)
    $inputsPanel.Controls.Add($txtUser)
    $inputsPanel.Controls.Add($chkIncludeDisabled)

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Multiline = $true
    $txtOutput.ReadOnly = $true
    $txtOutput.ScrollBars = "Both"
    $txtOutput.WordWrap = $false
    $txtOutput.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtOutput.Dock = "Fill"

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = "Preview"
    $btnPreview.AutoSize = $true

    $btnDryRun = New-Object System.Windows.Forms.Button
    $btnDryRun.Text = "Dry-Run Runner"
    $btnDryRun.AutoSize = $true

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = "Copy"
    $btnCopy.AutoSize = $true

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.AutoSize = $true

    $buttonsLeft = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonsLeft.FlowDirection = "LeftToRight"
    $buttonsLeft.WrapContents = $false
    $buttonsLeft.AutoSize = $true
    $buttonsLeft.Controls.Add($btnPreview)
    $buttonsLeft.Controls.Add($btnDryRun)
    $buttonsLeft.Dock = "Left"

    $buttonsRight = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonsRight.FlowDirection = "RightToLeft"
    $buttonsRight.WrapContents = $false
    $buttonsRight.AutoSize = $true
    $buttonsRight.Controls.Add($btnClose)
    $buttonsRight.Controls.Add($btnCopy)
    $buttonsRight.Dock = "Right"

    $buttonsRow = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonsRow.ColumnCount = 2
    $buttonsRow.Dock = "Fill"
    $buttonsRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $buttonsRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $buttonsRow.Controls.Add($buttonsLeft, 0, 0)
    $buttonsRow.Controls.Add($buttonsRight, 1, 0)

    $tlp.Controls.Add($inputsPanel, 0, 0)
    $tlp.SetColumnSpan($inputsPanel, 2)
    $tlp.Controls.Add($txtOutput, 0, 1)
    $tlp.SetColumnSpan($txtOutput, 2)
    $tlp.Controls.Add($buttonsRow, 0, 2)
    $tlp.SetColumnSpan($buttonsRow, 2)

    $dialog.Controls.Add($tlp)

    $btnClose.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $btnCopy.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($txtOutput.Text)
        [System.Windows.Forms.MessageBox]::Show("In Zwischenablage kopiert.", "Preview", "OK", "Information") | Out-Null
    })

    $btnPreview.Add_Click({
        $trigger = [string]$cmbTrigger.SelectedItem
        $userInput = $txtUser.Text.Trim()
        $includeDisabled = $chkIncludeDisabled.Checked

        $policy = Get-PreviewPolicyFromGrid -Rows $gridActionsPoc.Rows -IncludeDisabled:$includeDisabled
        $section = if ($trigger -eq "Logon") { $policy.logon } else { $policy.logoff }

        $lines = @()
        $lines += ("Trigger: {0} | User: {1}" -f $trigger, $userInput)
        $lines += "----------------------------------------"
        $lines += ""
        $lines += "[EVERY]"

        if ($section.every.enabled -and $section.every.actions) {
            foreach ($action in @($section.every.actions)) {
                $paramsText = if ($action.params) { ($action.params | ConvertTo-Json -Depth 6 -Compress) } else { "{}" }
                $lines += ("  [EVERY] - {0} (mode={1}, params={2})" -f $action.name, $action.mode, $paramsText)
            }
        } else {
            $lines += "  (keine)"
        }

        $lines += ""
        $lines += "[ONCE]"

        $stateRoot = Join-Path $env:LOCALAPPDATA ("CTX-Wartungs-Tools\\State\\{0}" -f $toolId)
        $onceItems = @($section.once)
        if ($onceItems.Count -eq 0) {
            $lines += "  (keine)"
        } else {
            foreach ($entry in $onceItems) {
                if (-not $entry) { continue }
                if (-not $entry.enabled -and -not $includeDisabled) { continue }
                $campaignId = [string]$entry.campaignId
                $validUntil = [string]$entry.validUntil
                $targets = $entry.targets

                $status = "WOULD RUN"
                $reason = ""
                if (-not $entry.enabled) {
                    $status = "SKIP"
                    $reason = "disabled"
                }

                if ($validUntil) {
                    try {
                        $until = [datetime]::Parse($validUntil)
                        if ($until.Date -lt (Get-Date).Date) {
                            $status = "SKIP"
                            $reason = "expired"
                        }
                    } catch {
                        $status = "SKIP"
                        $reason = "invalid validUntil"
                    }
                }

                $stateFile = Join-Path $stateRoot ("{0}_once_{1}.json" -f $trigger.ToLowerInvariant(), $campaignId)
                $doneAtText = $null
                if ($status -eq "WOULD RUN" -and $campaignId -and (Test-Path $stateFile)) {
                    $status = "SKIP"
                    $reason = "already done"
                    try {
                        $stateJson = Get-Content -Path $stateFile -Raw | ConvertFrom-Json
                        if ($stateJson.doneAt) {
                            $doneAtText = [string]$stateJson.doneAt
                        } elseif ($stateJson.doneAtUtc) {
                            $doneAtText = [string]$stateJson.doneAtUtc
                        }
                    } catch {
                        $doneAtText = $null
                    }
                }

                $targetsMatch = $true
                $groupsPresent = $false
                if ($targets) {
                    $groupsPresent = (@($targets.groups).Count -gt 0)
                    $targetsMatch = Test-UserTargetMatch -UserInput $userInput -Targets $targets
                }

                if ($status -eq "WOULD RUN") {
                    if ($groupsPresent) {
                        $status = "SKIP"
                        $reason = "groups not evaluated"
                    } elseif (-not $targetsMatch) {
                        $status = "SKIP"
                        $reason = "targets"
                    }
                }

                $statusLine = $status
                if ($status -eq "SKIP" -and $reason -eq "already done" -and $doneAtText) {
                    $statusLine = "$status (already done at $doneAtText)"
                } elseif ($reason) {
                    $statusLine = "$status ($reason)"
                }

                $validUntilText = if ($validUntil) { $validUntil } else { "" }
                $lines += ("  [ONCE] campaignId={0} validUntil={1}" -f $campaignId, $validUntilText)
                foreach ($action in @($entry.actions)) {
                    $paramsText = if ($action.params) { ($action.params | ConvertTo-Json -Depth 6 -Compress) } else { "{}" }
                    $lines += ("    [{0}] - {1} (mode={2}, params={3})" -f $statusLine, $action.name, $action.mode, $paramsText)
                }
            }
        }

        $txtOutput.Text = ($lines -join "`r`n")
    })

    $btnDryRun.Add_Click({
        $btnPreview.Enabled = $false
        $btnDryRun.Enabled = $false
        $btnCopy.Enabled = $false
        $btnClose.Enabled = $false

        try {
            $policyOut = Convert-PocRowsToPolicy -Rows $gridActionsPoc.Rows -BasePolicy (Get-DefaultPolicy)
            if ($null -eq $policyOut) { return }
            Write-PolicyJson -Policy $policyOut

            $trigger = [string]$cmbTrigger.SelectedItem
            $runnerPath = Join-Path $toolRoot "Runners\Runner.ps1"
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "pwsh"
            $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$runnerPath`" -Trigger $trigger -PreviewOnly"
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            [void]$process.Start()
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            $process.WaitForExit()

            $logRoot = Join-Path $toolRoot "logs"
            try {
                $customer = Get-CustomerConfig
                if ($customer.logging -and $customer.logging.relativeLogPath) {
                    $logRoot = Join-Path $toolRoot $customer.logging.relativeLogPath
                }
            } catch {
                $logRoot = Join-Path $toolRoot "logs"
            }
            $dateStamp = (Get-Date).ToString("yyyyMMdd")
            $logPath = Join-Path $logRoot ("$toolId-$dateStamp.log")

            $txtOutput.AppendText("`r`n----------------------------------------`r`n")
            $txtOutput.AppendText("DRY-RUN Runner output:`r`n")
            if ($stdout) { $txtOutput.AppendText($stdout + "`r`n") }
            if ($stderr) { $txtOutput.AppendText($stderr + "`r`n") }
            $txtOutput.AppendText("Log: $logPath`r`n")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Dry-run fehlgeschlagen: $($_.Exception.Message)", "Fehler", "OK", "Error") | Out-Null
        } finally {
            $btnPreview.Enabled = $true
            $btnDryRun.Enabled = $true
            $btnCopy.Enabled = $true
            $btnClose.Enabled = $true
        }
    })

    [void]$dialog.ShowDialog($form)
})

$chkPocShowEnabled.Add_CheckedChanged({
    foreach ($row in $gridActionsPoc.Rows) {
        if ($row.IsNewRow) { continue }
        if ($chkPocShowEnabled.Checked) {
            $row.Visible = [bool]$row.Cells[0].Value
        } else {
            $row.Visible = $true
        }
    }
})

function Get-PolicyPath {
    return (Join-Path $toolRoot "policy.json")
}

function Read-PolicyJson {
    $policyPath = Get-PolicyPath
    $policy = Get-DefaultPolicy

    if (Test-Path $policyPath) {
        try {
            $loaded = Get-Content $policyPath -Raw | ConvertFrom-Json
            if ($loaded.logon) {
                if ($loaded.logon.every) { $policy.logon.every = $loaded.logon.every }
                if ($loaded.logon.once) { $policy.logon.once = $loaded.logon.once }
            }
            if ($loaded.logoff) {
                if ($loaded.logoff.every) { $policy.logoff.every = $loaded.logoff.every }
                if ($loaded.logoff.once) { $policy.logoff.once = $loaded.logoff.once }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("policy.json ist ungueltig. Defaults geladen.", "Hinweis", "OK", "Warning") | Out-Null
            $policy = Get-DefaultPolicy
        }
    }

    return $policy
}

function Write-PolicyJson {
    param(
        [Parameter(Mandatory)]
        [object]$Policy
    )

    $policyPath = Get-PolicyPath
    $json = $Policy | ConvertTo-Json -Depth 8
    Set-Content -Path $policyPath -Value $json -Encoding UTF8
}

function Add-PocRow {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [hashtable]$RowData
    )

    $rowIndex = $Grid.Rows.Add()
    $row = $Grid.Rows[$rowIndex]

    $row.Cells[0].Value = [bool]$RowData.Enabled
    $row.Cells[1].Value = $RowData.Trigger
    $row.Cells[2].Value = $RowData.Frequency
    $row.Cells[3].Value = $RowData.Action
    $row.Cells[4].Value = $RowData.Mode
    $row.Cells[5].Value = $RowData.Params
    $row.Cells[6].Value = $RowData.CampaignId
    $row.Cells[7].Value = $RowData.ValidUntil
    $row.Cells[8].Value = $RowData.TargetsUsers
    $row.Cells[9].Value = $RowData.TargetsGroups
}

function Convert-PolicyToPocRows {
    param(
        [object]$Policy
    )

    $gridActionsPoc.DataSource = $null
    $gridActionsPoc.Rows.Clear()

    $logonEvery = Normalize-Every $Policy.logon.every
    $logoffEvery = Normalize-Every $Policy.logoff.every

    if ($logonEvery -and $logonEvery.actions) {
        foreach ($action in @($logonEvery.actions)) {
            Add-PocRow -Grid $gridActionsPoc -RowData @{
                Enabled = [bool]$logonEvery.enabled
                Trigger = "Logon"
                Frequency = "Every"
                Action = $action.name
                Mode = if ($action.mode) { $action.mode } else { "Silent" }
                Params = if ($action.params) { ($action.params | ConvertTo-Json -Depth 6 -Compress) } else { "" }
                CampaignId = ""
                ValidUntil = ""
                TargetsUsers = ""
                TargetsGroups = ""
            }
        }
    }

    if ($logoffEvery -and $logoffEvery.actions) {
        foreach ($action in @($logoffEvery.actions)) {
            Add-PocRow -Grid $gridActionsPoc -RowData @{
                Enabled = [bool]$logoffEvery.enabled
                Trigger = "Logoff"
                Frequency = "Every"
                Action = $action.name
                Mode = if ($action.mode) { $action.mode } else { "Silent" }
                Params = if ($action.params) { ($action.params | ConvertTo-Json -Depth 6 -Compress) } else { "" }
                CampaignId = ""
                ValidUntil = ""
                TargetsUsers = ""
                TargetsGroups = ""
            }
        }
    }

    foreach ($entry in @($Policy.logon.once)) {
        if (-not $entry) { continue }
        $targets = $entry.targets
        $users = if ($targets) { @($targets.users) } else { @() }
        $groups = if ($targets) { @($targets.groups) } else { @() }
        foreach ($action in @($entry.actions)) {
            Add-PocRow -Grid $gridActionsPoc -RowData @{
                Enabled = [bool]$entry.enabled
                Trigger = "Logon"
                Frequency = "Once"
                Action = $action.name
                Mode = if ($action.mode) { $action.mode } else { "Silent" }
                Params = if ($action.params) { ($action.params | ConvertTo-Json -Depth 6 -Compress) } else { "" }
                CampaignId = $entry.campaignId
                ValidUntil = $entry.validUntil
                TargetsUsers = ($users -join "`r`n")
                TargetsGroups = ($groups -join "`r`n")
            }
        }
    }

    foreach ($entry in @($Policy.logoff.once)) {
        if (-not $entry) { continue }
        $targets = $entry.targets
        $users = if ($targets) { @($targets.users) } else { @() }
        $groups = if ($targets) { @($targets.groups) } else { @() }
        foreach ($action in @($entry.actions)) {
            Add-PocRow -Grid $gridActionsPoc -RowData @{
                Enabled = [bool]$entry.enabled
                Trigger = "Logoff"
                Frequency = "Once"
                Action = $action.name
                Mode = if ($action.mode) { $action.mode } else { "Silent" }
                Params = if ($action.params) { ($action.params | ConvertTo-Json -Depth 6 -Compress) } else { "" }
                CampaignId = $entry.campaignId
                ValidUntil = $entry.validUntil
                TargetsUsers = ($users -join "`r`n")
                TargetsGroups = ($groups -join "`r`n")
            }
        }
    }
}

function Split-TargetText {
    param([string]$Text)

    if (-not $Text) { return @() }
    return $Text -split "[,\r\n]" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Convert-PocRowsToPolicy {
    param(
        [System.Windows.Forms.DataGridViewRowCollection]$Rows,
        [object]$BasePolicy
    )

    $policy = $BasePolicy
    $policy.logon.every.actions = @()
    $policy.logoff.every.actions = @()
    $policy.logon.once = @()
    $policy.logoff.once = @()

    $logonEveryActions = @()
    $logoffEveryActions = @()
    $logonOnceGroups = @{}
    $logoffOnceGroups = @{}

    $rowNumber = 1
    foreach ($row in $Rows) {
        if ($row.IsNewRow) { continue }
        $enabled = [bool]$row.Cells[0].Value
        if (-not $enabled) { continue }

        $trigger = [string]$row.Cells[1].Value
        $frequency = [string]$row.Cells[2].Value
        $actionName = [string]$row.Cells[3].Value
        $mode = [string]$row.Cells[4].Value
        $paramsText = [string]$row.Cells[5].Value
        $campaignId = [string]$row.Cells[6].Value
        $validUntil = [string]$row.Cells[7].Value
        $targetsUsersText = [string]$row.Cells[8].Value
        $targetsGroupsText = [string]$row.Cells[9].Value

        if (-not $actionName) { continue }
        if (-not $mode) { $mode = "Silent" }

        $params = @{}
        if ([string]::IsNullOrWhiteSpace($paramsText)) {
            $params = @{}
        } else {
            try {
                $parsed = $paramsText | ConvertFrom-Json -ErrorAction Stop
                if ($parsed -is [System.Collections.IDictionary] -or $parsed -is [pscustomobject]) {
                    $params = $parsed
                } else {
                    $params = @{}
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Ungültiges JSON in Params (Zeile $rowNumber). Bitte korrigieren.", "Validierung", "OK", "Warning") | Out-Null
                return $null
            }
        }

        if ($null -eq $users) { $users = @() }
        if ($null -eq $groups) { $groups = @() }

        $actionObj = [pscustomobject]@{
            name = $actionName
            mode = $mode
            params = $params
        }

        if ($frequency -eq "Every") {
            if ($trigger -eq "Logon" -or $trigger -eq "Both") {
                $logonEveryActions += $actionObj
            }
            if ($trigger -eq "Logoff" -or $trigger -eq "Both") {
                $logoffEveryActions += $actionObj
            }
            $rowNumber += 1
            continue
        }

        if ($frequency -ne "Once") {
            $rowNumber += 1
            continue
        }

        $users = Split-TargetText -Text $targetsUsersText
        $groups = Split-TargetText -Text $targetsGroupsText
        $key = ($campaignId + "|" + $validUntil + "|" + ($users -join ";") + "|" + ($groups -join ";"))

        if ($trigger -eq "Logon" -or $trigger -eq "Both") {
            if (-not $logonOnceGroups.ContainsKey($key)) {
                $logonOnceGroups[$key] = [pscustomobject]@{
                    enabled = $true
                    campaignId = $campaignId
                    validUntil = $validUntil
                    targets = [pscustomobject]@{ users = $users; groups = $groups }
                    actions = @()
                }
            }
            $logonOnceGroups[$key].actions += $actionObj
        }

        if ($trigger -eq "Logoff" -or $trigger -eq "Both") {
            if (-not $logoffOnceGroups.ContainsKey($key)) {
                $logoffOnceGroups[$key] = [pscustomobject]@{
                    enabled = $true
                    campaignId = $campaignId
                    validUntil = $validUntil
                    targets = [pscustomobject]@{ users = $users; groups = $groups }
                    actions = @()
                }
            }
            $logoffOnceGroups[$key].actions += $actionObj
        }

        $rowNumber += 1
    }

    $policy.logon.every.actions = $logonEveryActions
    $policy.logoff.every.actions = $logoffEveryActions
    $policy.logon.every.enabled = ($logonEveryActions.Count -gt 0)
    $policy.logoff.every.enabled = ($logoffEveryActions.Count -gt 0)
    $policy.logon.once = @($logonOnceGroups.Values)
    $policy.logoff.once = @($logoffOnceGroups.Values)

    return $policy
}

function Load-PolicyIntoUI {
    $policy = Read-PolicyJson

    $logonEvery = Normalize-Every $policy.logon.every
    $logoffEvery = Normalize-Every $policy.logoff.every
    $logonOnce = @(@($policy.logon.once) | Where-Object { $_ })
    $logoffOnce = @(@($policy.logoff.once) | Where-Object { $_ })

    $chkLogonEveryEnabled.Checked = $false
    $chkLogoffEveryEnabled.Checked = $false

    if ($logonEvery) {
        $chkLogonEveryEnabled.Checked = [bool]$logonEvery.enabled
        $logonEveryActionsList = Normalize-Actions $logonEvery.actions
        Set-ActionsToGrid -Grid $logonEveryActions.Grid -ActionNames $actionNames -Actions $logonEveryActionsList
    } else {
        $logonEveryActions.Grid.Rows.Clear()
    }

    if ($logoffEvery) {
        $chkLogoffEveryEnabled.Checked = [bool]$logoffEvery.enabled
        $logoffEveryActionsList = Normalize-Actions $logoffEvery.actions
        Set-ActionsToGrid -Grid $logoffEveryActions.Grid -ActionNames $actionNames -Actions $logoffEveryActionsList
    } else {
        $logoffEveryActions.Grid.Rows.Clear()
    }

    $logonOnceItem = if ($logonOnce.Count -gt 0) { $logonOnce[0] } else { $null }
    $logoffOnceItem = if ($logoffOnce.Count -gt 0) { $logoffOnce[0] } else { $null }

    $chkLogonOnceEnabled.Checked = $false
    $logonOnceCampaign.TextBox.Text = ""
    $logonOnceValidUntil.TextBox.Text = ""
    $logonOnceUsers.TextBox.Text = ""
    $logonOnceGroups.TextBox.Text = ""
    $logonOnceActions.Grid.Rows.Clear()

    if ($logonOnceItem) {
        $chkLogonOnceEnabled.Checked = [bool]$logonOnceItem.enabled
        $logonOnceCampaign.TextBox.Text = [string]$logonOnceItem.campaignId
        $logonOnceValidUntil.TextBox.Text = [string]$logonOnceItem.validUntil
        $logonOnceTargets = $logonOnceItem.targets
        $logonOnceUsersList = if ($logonOnceTargets) { Get-TargetLines -Value $logonOnceTargets.users } else { @() }
        $logonOnceGroupsList = if ($logonOnceTargets) { Get-TargetLines -Value $logonOnceTargets.groups } else { @() }
        $logonOnceUsers.TextBox.Text = ($logonOnceUsersList -join "`r`n")
        $logonOnceGroups.TextBox.Text = ($logonOnceGroupsList -join "`r`n")
        $logonOnceActionsList = Normalize-Actions $logonOnceItem.actions
        Set-ActionsToGrid -Grid $logonOnceActions.Grid -ActionNames $actionNames -Actions $logonOnceActionsList
    }

    $chkLogoffOnceEnabled.Checked = $false
    $logoffOnceCampaign.TextBox.Text = ""
    $logoffOnceValidUntil.TextBox.Text = ""
    $logoffOnceUsers.TextBox.Text = ""
    $logoffOnceGroups.TextBox.Text = ""
    $logoffOnceActions.Grid.Rows.Clear()

    if ($logoffOnceItem) {
        $chkLogoffOnceEnabled.Checked = [bool]$logoffOnceItem.enabled
        $logoffOnceCampaign.TextBox.Text = [string]$logoffOnceItem.campaignId
        $logoffOnceValidUntil.TextBox.Text = [string]$logoffOnceItem.validUntil
        $logoffOnceTargets = $logoffOnceItem.targets
        $logoffOnceUsersList = if ($logoffOnceTargets) { Get-TargetLines -Value $logoffOnceTargets.users } else { @() }
        $logoffOnceGroupsList = if ($logoffOnceTargets) { Get-TargetLines -Value $logoffOnceTargets.groups } else { @() }
        $logoffOnceUsers.TextBox.Text = ($logoffOnceUsersList -join "`r`n")
        $logoffOnceGroups.TextBox.Text = ($logoffOnceGroupsList -join "`r`n")
        $logoffOnceActionsList = Normalize-Actions $logoffOnceItem.actions
        Set-ActionsToGrid -Grid $logoffOnceActions.Grid -ActionNames $actionNames -Actions $logoffOnceActionsList
    }

    $lblPolicyStatus.Text = "Policy geladen."

    if ($logonOnce.Count -gt 1 -or $logoffOnce.Count -gt 1) {
        $lblPolicyStatus.Text = "Policy geladen (mehrere Once-Kampagnen gefunden; nur erste angezeigt)."
    }
}

function Build-OnceEntry {
    param(
        [bool]$Enabled,
        [string]$CampaignId,
        [string]$ValidUntil,
        [string[]]$Users,
        [string[]]$Groups,
        [object[]]$Actions
    )

    return [pscustomobject]@{
        enabled = $Enabled
        campaignId = $CampaignId
        validUntil = $ValidUntil
        targets = [pscustomobject]@{
            users = $Users
            groups = $Groups
        }
        actions = $Actions
    }
}

function Save-UIToPolicy {
    $logonEveryActionsList = Get-ActionsFromGrid -Grid $logonEveryActions.Grid
    $logoffEveryActionsList = Get-ActionsFromGrid -Grid $logoffEveryActions.Grid

    if ($null -eq $logonEveryActionsList) { $logonEveryActionsList = @() }
    if ($null -eq $logoffEveryActionsList) { $logoffEveryActionsList = @() }

    if ($logonEveryActionsList -isnot [System.Array]) { $logonEveryActionsList = @($logonEveryActionsList) }
    if ($logoffEveryActionsList -isnot [System.Array]) { $logoffEveryActionsList = @($logoffEveryActionsList) }

    $logonOnceActionsList = Get-ActionsFromGrid -Grid $logonOnceActions.Grid
    $logoffOnceActionsList = Get-ActionsFromGrid -Grid $logoffOnceActions.Grid

    $logonOnceCampaignId = $logonOnceCampaign.TextBox.Text.Trim()
    $logoffOnceCampaignId = $logoffOnceCampaign.TextBox.Text.Trim()

    $logonOnceValidUntilValue = $logonOnceValidUntil.TextBox.Text.Trim()
    $logoffOnceValidUntilValue = $logoffOnceValidUntil.TextBox.Text.Trim()

    if ($logonOnceValidUntilValue -and ($logonOnceValidUntilValue -notmatch "^\d{4}-\d{2}-\d{2}$")) {
        [System.Windows.Forms.MessageBox]::Show("logon.once: ValidUntil muss YYYY-MM-DD sein.", "Validierung", "OK", "Warning") | Out-Null
        return
    }

    if ($logoffOnceValidUntilValue -and ($logoffOnceValidUntilValue -notmatch "^\d{4}-\d{2}-\d{2}$")) {
        [System.Windows.Forms.MessageBox]::Show("logoff.once: ValidUntil muss YYYY-MM-DD sein.", "Validierung", "OK", "Warning") | Out-Null
        return
    }

    $logonOnceUsersList = Get-Lines -Text $logonOnceUsers.TextBox.Text
    $logonOnceGroupsList = Get-Lines -Text $logonOnceGroups.TextBox.Text
    $logoffOnceUsersList = Get-Lines -Text $logoffOnceUsers.TextBox.Text
    $logoffOnceGroupsList = Get-Lines -Text $logoffOnceGroups.TextBox.Text

    if ($null -eq $logonOnceUsersList) { $logonOnceUsersList = @() }
    if ($null -eq $logonOnceGroupsList) { $logonOnceGroupsList = @() }
    if ($null -eq $logoffOnceUsersList) { $logoffOnceUsersList = @() }
    if ($null -eq $logoffOnceGroupsList) { $logoffOnceGroupsList = @() }

    $hasLogonOnceTargets = ($logonOnceUsersList.Count -gt 0) -or ($logonOnceGroupsList.Count -gt 0)
    $hasLogoffOnceTargets = ($logoffOnceUsersList.Count -gt 0) -or ($logoffOnceGroupsList.Count -gt 0)

    $hasLogonOnceInput = ($logonOnceCampaignId -or $logonOnceActionsList.Count -gt 0 -or $hasLogonOnceTargets)
    $hasLogoffOnceInput = ($logoffOnceCampaignId -or $logoffOnceActionsList.Count -gt 0 -or $hasLogoffOnceTargets)

    if ($hasLogonOnceInput) {
        if (-not $logonOnceCampaignId) {
            [System.Windows.Forms.MessageBox]::Show("logon.once: CampaignId darf nicht leer sein.", "Validierung", "OK", "Warning") | Out-Null
            return
        }
    }

    if ($hasLogoffOnceInput) {
        if (-not $logoffOnceCampaignId) {
            [System.Windows.Forms.MessageBox]::Show("logoff.once: CampaignId darf nicht leer sein.", "Validierung", "OK", "Warning") | Out-Null
            return
        }
    }

    $logonOnceEntry = $null
    $logoffOnceEntry = $null

    if ($hasLogonOnceInput) {
        $logonOnceEntry = Build-OnceEntry -Enabled $chkLogonOnceEnabled.Checked -CampaignId $logonOnceCampaignId -ValidUntil $logonOnceValidUntilValue -Users $logonOnceUsersList -Groups $logonOnceGroupsList -Actions $logonOnceActionsList
    }

    if ($hasLogoffOnceInput) {
        $logoffOnceEntry = Build-OnceEntry -Enabled $chkLogoffOnceEnabled.Checked -CampaignId $logoffOnceCampaignId -ValidUntil $logoffOnceValidUntilValue -Users $logoffOnceUsersList -Groups $logoffOnceGroupsList -Actions $logoffOnceActionsList
    }

    $policyOut = [pscustomobject]@{
        logon = [pscustomobject]@{
            every = [pscustomobject]@{
                enabled = $chkLogonEveryEnabled.Checked
                actions = $logonEveryActionsList
            }
            once = @()
        }
        logoff = [pscustomobject]@{
            every = [pscustomobject]@{
                enabled = $chkLogoffEveryEnabled.Checked
                actions = $logoffEveryActionsList
            }
            once = @()
        }
    }

    if ($logonOnceEntry) { $policyOut.logon.once = @($logonOnceEntry) }
    if ($logoffOnceEntry) { $policyOut.logoff.once = @($logoffOnceEntry) }

    return $policyOut
}

$btnPolicyLoad.Add_Click({
    try {
        Load-PolicyIntoUI
        Convert-PolicyToPocRows -Policy (Read-PolicyJson)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Policy konnte nicht geladen werden: $($_.Exception.Message)", "Fehler", "OK", "Error") | Out-Null
    }
})

$btnPolicySave.Add_Click({
    try {
        $policyOut = Convert-PocRowsToPolicy -Rows $gridActionsPoc.Rows -BasePolicy (Get-DefaultPolicy)
        if ($null -eq $policyOut) { return }

        Write-PolicyJson -Policy $policyOut
        $lblPolicyStatus.Text = "Policy gespeichert: $(Get-PolicyPath)"
        Write-UiLog "Policy gespeichert" "INFO"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Policy konnte nicht gespeichert werden: $($_.Exception.Message)", "Fehler", "OK", "Error") | Out-Null
    }
})

$cfgPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$cfgPanel.Dock = "Fill"
$cfgPanel.FlowDirection = "TopDown"
$cfgPanel.WrapContents = $false
$cfgPanel.AutoScroll = $true

$lblConfigWarning = New-Object System.Windows.Forms.Label
$lblConfigWarning.AutoSize = $true
$lblConfigWarning.ForeColor = [System.Drawing.Color]::Red
$lblConfigWarning.Text = ""

$cfgCustomerName = New-LabeledTextBox -LabelText "Customer Name" -TextWidth 260
$cfgRepoRoot = New-LabeledTextBox -LabelText "Repo Root (optional)" -TextWidth 260
$cfgWindowTitle = New-LabeledTextBox -LabelText "Window Title" -TextWidth 260
$cfgSupportText = New-LabeledTextBox -LabelText "Support Text" -TextWidth 260
$cfgRelativeLog = New-LabeledTextBox -LabelText "Relative Log Path" -TextWidth 180
$cfgAdminLogRoot = New-LabeledTextBox -LabelText "Admin Log Root" -TextWidth 260
$cfgFslogixProfile = New-LabeledTextBox -LabelText "FSLogix ProfileShare" -TextWidth 260
$cfgFslogixOffice = New-LabeledTextBox -LabelText "FSLogix OfficeContainerShare" -TextWidth 260

$chkAllowOffline = New-Object System.Windows.Forms.CheckBox
$chkAllowOffline.Text = "Allow Offline"

$chkAllowLogoff = New-Object System.Windows.Forms.CheckBox
$chkAllowLogoff.Text = "Allow Logoff Runner"

$chkFslogixEnabled = New-Object System.Windows.Forms.CheckBox
$chkFslogixEnabled.Text = "FSLogix Enabled"

$cfgButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$cfgButtons.FlowDirection = "LeftToRight"
$cfgButtons.WrapContents = $false
$cfgButtons.AutoSize = $true

$btnConfigNew = New-Object System.Windows.Forms.Button
$btnConfigNew.Text = "Neu anlegen (Defaults)"
$btnConfigNew.AutoSize = $true

$btnConfigSave = New-Object System.Windows.Forms.Button
$btnConfigSave.Text = "Speichern"
$btnConfigSave.AutoSize = $true

$btnConfigOpen = New-Object System.Windows.Forms.Button
$btnConfigOpen.Text = "customer.json oeffnen"
$btnConfigOpen.AutoSize = $true

$btnOpenLogs = New-Object System.Windows.Forms.Button
$btnOpenLogs.Text = "Logs-Ordner oeffnen"
$btnOpenLogs.AutoSize = $true

$cfgButtons.Controls.Add($btnConfigNew)
$cfgButtons.Controls.Add($btnConfigSave)
$cfgButtons.Controls.Add($btnConfigOpen)
$cfgButtons.Controls.Add($btnOpenLogs)

$lblConfigStatus = New-Object System.Windows.Forms.Label
$lblConfigStatus.AutoSize = $true
$lblConfigStatus.Text = ""

$cfgPanel.Controls.Add($lblConfigWarning)
$cfgPanel.Controls.Add($cfgCustomerName.Panel)
$cfgPanel.Controls.Add($cfgRepoRoot.Panel)
$cfgPanel.Controls.Add($cfgWindowTitle.Panel)
$cfgPanel.Controls.Add($cfgSupportText.Panel)
$cfgPanel.Controls.Add($cfgRelativeLog.Panel)
$cfgPanel.Controls.Add($cfgAdminLogRoot.Panel)
$cfgPanel.Controls.Add($chkAllowOffline)
$cfgPanel.Controls.Add($chkAllowLogoff)
$cfgPanel.Controls.Add($chkFslogixEnabled)
$cfgPanel.Controls.Add($cfgFslogixProfile.Panel)
$cfgPanel.Controls.Add($cfgFslogixOffice.Panel)
$cfgPanel.Controls.Add($cfgButtons)
$cfgPanel.Controls.Add($lblConfigStatus)

$tabConfig.Controls.Add($cfgPanel)

function Apply-CustomerToUi {
    param([pscustomobject]$Customer)

    $cfgCustomerName.TextBox.Text = [string]$Customer.customer.name
    $cfgRepoRoot.TextBox.Text = [string]$Customer.paths.repoRoot
    $cfgWindowTitle.TextBox.Text = [string]$Customer.branding.windowTitle
    $cfgSupportText.TextBox.Text = [string]$Customer.branding.supportText
    $cfgRelativeLog.TextBox.Text = [string]$Customer.logging.relativeLogPath
    $cfgAdminLogRoot.TextBox.Text = [string]$Customer.logging.adminLogRoot
    $chkAllowOffline.Checked = [bool]$Customer.flags.allowOffline
    $chkAllowLogoff.Checked = [bool]$Customer.flags.allowLogoffRunner
    $chkFslogixEnabled.Checked = [bool]$Customer.fslogix.enabled
    $cfgFslogixProfile.TextBox.Text = [string]$Customer.fslogix.profileShare
    $cfgFslogixOffice.TextBox.Text = [string]$Customer.fslogix.officeContainerShare
}

function Load-CustomerIntoUi {
    $path = Join-Path $toolRoot "customer.json"
    $lblConfigWarning.Text = ""

    if (-not (Test-Path $path)) {
        $lblConfigWarning.Text = "customer.json fehlt. Bitte neu anlegen."
        Apply-CustomerToUi (Get-DefaultCustomer)
        return
    }

    try {
        $customer = Get-Content $path -Raw | ConvertFrom-Json
        Apply-CustomerToUi $customer
    } catch {
        $lblConfigWarning.Text = "customer.json ist ungueltig. Bitte neu anlegen."
        Apply-CustomerToUi (Get-DefaultCustomer)
    }
}

function Save-CustomerFromUi {
    $path = Join-Path $toolRoot "customer.json"

    $name = $cfgCustomerName.TextBox.Text.Trim()
    $title = $cfgWindowTitle.TextBox.Text.Trim()
    $relLog = $cfgRelativeLog.TextBox.Text.Trim()

    if (-not $name) {
        [System.Windows.Forms.MessageBox]::Show("Customer Name ist Pflicht.", "Validierung", "OK", "Warning") | Out-Null
        return
    }

    if (-not $title) {
        [System.Windows.Forms.MessageBox]::Show("Window Title ist Pflicht.", "Validierung", "OK", "Warning") | Out-Null
        return
    }

    if (-not $relLog) {
        [System.Windows.Forms.MessageBox]::Show("Relative Log Path ist Pflicht.", "Validierung", "OK", "Warning") | Out-Null
        return
    }

    if ([System.IO.Path]::IsPathRooted($relLog)) {
        [System.Windows.Forms.MessageBox]::Show("Relative Log Path darf kein absoluter Pfad sein.", "Validierung", "OK", "Warning") | Out-Null
        return
    }

    $customer = [pscustomobject]@{
        customer = [pscustomobject]@{ name = $name }
        paths = [pscustomobject]@{ repoRoot = $cfgRepoRoot.TextBox.Text.Trim() }
        fslogix = [pscustomobject]@{
            enabled = $chkFslogixEnabled.Checked
            profileShare = $cfgFslogixProfile.TextBox.Text.Trim()
            officeContainerShare = $cfgFslogixOffice.TextBox.Text.Trim()
        }
        branding = [pscustomobject]@{
            windowTitle = $title
            supportText = $cfgSupportText.TextBox.Text.Trim()
        }
        logging = [pscustomobject]@{
            relativeLogPath = $relLog
            adminLogRoot = $cfgAdminLogRoot.TextBox.Text.Trim()
        }
        flags = [pscustomobject]@{
            allowOffline = $chkAllowOffline.Checked
            allowLogoffRunner = $chkAllowLogoff.Checked
        }
    }

    $json = $customer | ConvertTo-Json -Depth 6
    Set-Content -Path $path -Value $json -Encoding UTF8

    $lblConfigWarning.Text = ""
    $lblConfigStatus.Text = "Gespeichert: $path"
    $form.Text = "Wartung Admin - $name"
}

$btnConfigNew.Add_Click({
    Apply-CustomerToUi (Get-DefaultCustomer)
    $lblConfigWarning.Text = "Defaults geladen. Bitte speichern."
})

$btnConfigSave.Add_Click({
    try {
        Save-CustomerFromUi
    } catch {
        [System.Windows.Forms.MessageBox]::Show("customer.json konnte nicht gespeichert werden: $($_.Exception.Message)", "Fehler", "OK", "Error") | Out-Null
    }
})

$btnConfigOpen.Add_Click({
    $path = Join-Path $toolRoot "customer.json"
    if (Test-Path $path) {
        Start-Process $path
    }
})

$btnOpenLogs.Add_Click({
    $path = Join-Path $toolRoot "customer.json"
    $relLog = "logs"
    if (Test-Path $path) {
        try {
            $customer = Get-Content $path -Raw | ConvertFrom-Json
            if ($customer.logging.relativeLogPath) {
                $relLog = $customer.logging.relativeLogPath
            }
        } catch {
            $relLog = "logs"
        }
    }
    $logRoot = Join-Path $toolRoot $relLog
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    Start-Process explorer.exe "`"$logRoot`""
})

Load-CustomerIntoUi
if ($autoLoadPolicy) {
    Load-PolicyIntoUI
    Convert-PolicyToPocRows -Policy (Read-PolicyJson)
}

$form.Controls.Add($tabs)
[void]$form.ShowDialog()
Write-UiLog "Admin GUI geschlossen" "INFO"
