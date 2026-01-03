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

$tlpPolicy = New-Object System.Windows.Forms.TableLayoutPanel
$tlpPolicy.Dock = "Fill"
$tlpPolicy.ColumnCount = 2
$tlpPolicy.RowCount = 2
$tlpPolicy.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tlpPolicy.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tlpPolicy.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tlpPolicy.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

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

$tlpPolicy.Controls.Add($grpLogonEvery, 0, 0)
$tlpPolicy.Controls.Add($grpLogonOnce, 0, 1)
$tlpPolicy.Controls.Add($grpLogoffEvery, 1, 0)
$tlpPolicy.Controls.Add($grpLogoffOnce, 1, 1)

$policyRoot.Controls.Add($tlpPolicy, 0, 0)

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

function Load-PolicyIntoUi {
    $policyPath = Join-Path $toolRoot "policy.json"
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

function Save-PolicyFromUi {
    $policyPath = Join-Path $toolRoot "policy.json"

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

    $json = $policyOut | ConvertTo-Json -Depth 8
    Set-Content -Path $policyPath -Value $json -Encoding UTF8

    $lblPolicyStatus.Text = "Policy gespeichert: $policyPath"
    Write-UiLog "Policy gespeichert" "INFO"
}

$btnPolicyLoad.Add_Click({
    try {
        Load-PolicyIntoUi
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Policy konnte nicht geladen werden: $($_.Exception.Message)", "Fehler", "OK", "Error") | Out-Null
    }
})

$btnPolicySave.Add_Click({
    try {
        Save-PolicyFromUi
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
    Load-PolicyIntoUi
}

$form.Controls.Add($tabs)
[void]$form.ShowDialog()
Write-UiLog "Admin GUI geschlossen" "INFO"
