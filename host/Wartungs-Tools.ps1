param(
    [ValidateSet("User","Admin")]
    [string]$UiMode = "User"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$hostRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $hostRoot
$toolsRoot = Join-Path $repoRoot "tools"

# Discover all tools
$tools = @()
if (Test-Path $toolsRoot) {
    foreach ($toolDir in Get-ChildItem -Path $toolsRoot -Directory) {
        $manifestPath = Join-Path $toolDir.FullName "tool.json"
        if (Test-Path $manifestPath) {
            try {
                $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
                $tools += [pscustomobject]@{
                    Name = $manifest.name
                    ToolId = $manifest.toolId
                    Version = $manifest.version
                    Description = $manifest.description
                    EntryPoint = $manifest.entryPoint
                    ToolRoot = $toolDir.FullName
                }
            } catch { }
        }
    }
}

# Build GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "CTX Wartungs-Tools"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.MinimumSize = New-Object System.Drawing.Size(500, 300)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = "Fill"
$layout.RowCount = 3
$layout.ColumnCount = 1
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($layout)

# Header
$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text = "CTX Wartungs-Tools"
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblHeader.AutoSize = $true
$lblHeader.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 5)
$layout.Controls.Add($lblHeader, 0, 0)

# Tool list
$listView = New-Object System.Windows.Forms.ListView
$listView.View = "Details"
$listView.FullRowSelect = $true
$listView.MultiSelect = $false
$listView.Dock = "Fill"
$listView.GridLines = $true
[void]$listView.Columns.Add("Tool", 180)
[void]$listView.Columns.Add("Version", 70)
[void]$listView.Columns.Add("Beschreibung", 300)

foreach ($tool in $tools) {
    $item = New-Object System.Windows.Forms.ListViewItem($tool.Name)
    [void]$item.SubItems.Add($tool.Version)
    [void]$item.SubItems.Add($tool.Description)
    $item.Tag = $tool
    [void]$listView.Items.Add($item)
}

if ($listView.Items.Count -gt 0) {
    $listView.Items[0].Selected = $true
}

$layout.Controls.Add($listView, 0, 1)

# Buttons
$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$btnPanel.FlowDirection = "LeftToRight"
$btnPanel.AutoSize = $true
$btnPanel.Dock = "Fill"
$btnPanel.Padding = New-Object System.Windows.Forms.Padding(7, 5, 7, 5)

$btnUser = New-Object System.Windows.Forms.Button
$btnUser.Text = "User GUI starten"
$btnUser.AutoSize = $true
$btnUser.Enabled = ($listView.Items.Count -gt 0)

$btnAdmin = New-Object System.Windows.Forms.Button
$btnAdmin.Text = "Admin GUI starten"
$btnAdmin.AutoSize = $true
$btnAdmin.Enabled = ($listView.Items.Count -gt 0)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Schliessen"
$btnClose.AutoSize = $true

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.AutoSize = $true
$lblInfo.Padding = New-Object System.Windows.Forms.Padding(15, 8, 0, 0)
$lblInfo.Text = ("{0} Tool(s) gefunden" -f $tools.Count)

$btnPanel.Controls.Add($btnUser)
$btnPanel.Controls.Add($btnAdmin)
$btnPanel.Controls.Add($btnClose)
$btnPanel.Controls.Add($lblInfo)
$layout.Controls.Add($btnPanel, 0, 2)

# Helper to get selected tool
function Get-SelectedTool {
    if ($listView.SelectedItems.Count -eq 0) { return $null }
    return $listView.SelectedItems[0].Tag
}

function Start-Tool {
    param([string]$Mode)
    $tool = Get-SelectedTool
    if (-not $tool) {
        [System.Windows.Forms.MessageBox]::Show("Bitte ein Tool auswaehlen.", "Hinweis", "OK", "Information") | Out-Null
        return
    }
    $entryPoint = Join-Path $tool.ToolRoot $tool.EntryPoint
    if (-not (Test-Path $entryPoint)) {
        [System.Windows.Forms.MessageBox]::Show("EntryPoint nicht gefunden: $entryPoint", "Fehler", "OK", "Error") | Out-Null
        return
    }
    $lblInfo.Text = ("Starte {0} ({1})..." -f $tool.Name, $Mode)
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$entryPoint`" -UiMode $Mode"
}

$btnUser.Add_Click({ Start-Tool -Mode "User" })
$btnAdmin.Add_Click({ Start-Tool -Mode "Admin" })
$btnClose.Add_Click({ $form.Close() })

$listView.Add_DoubleClick({
    $mode = if ($UiMode -eq "Admin") { "Admin" } else { "User" }
    Start-Tool -Mode $mode
})

[void]$form.ShowDialog()
