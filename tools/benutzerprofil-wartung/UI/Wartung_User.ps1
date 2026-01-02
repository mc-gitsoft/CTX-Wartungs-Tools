param(
    [ValidateSet("User","Admin")]
    [string]$UiMode = "User"
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$modulePath = Join-Path $toolRoot "shared\WartungsTools.SDK.psm1"
Import-Module $modulePath -Force

$toolManifest = Get-Content (Join-Path $toolRoot "tool.json") -Raw | ConvertFrom-Json
$toolId = $toolManifest.toolId

try {
    $customer = Get-CustomerConfig
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "customer.json fehlt oder ist ungueltig. Bitte customer.json.example kopieren.",
        "Konfigurationsfehler",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

$relativeLogPath = "logs"
if ($customer.logging -and $customer.logging.relativeLogPath) {
    $relativeLogPath = [string]$customer.logging.relativeLogPath
}
$logRoot = Join-Path $toolRoot $relativeLogPath

function Write-UiLog {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO",
        [string]$Action
    )

    Write-Log -Level $Level -Message $Message -ToolId $toolId -Trigger "UserUI" -Action $Action -LogRoot $logRoot
}

$sessionId = (Get-Date).ToString("yyyyMMdd_HHmmss")
Write-UiLog "UI gestartet (Session=$sessionId)" "INFO"

$actions = @(
    @{
        Id = "Teams_Reset"
        Name = "Teams zuruecksetzen"
        Action = "Teams_Reset"
        Category = "Kommunikation"
        ConfirmMessage = ""
        Description = "Bereinigt Microsoft Teams Cache und Konfigurationsdateien im Benutzerprofil."
        Params = @{}
    },
    @{
        Id = "WorkspaceApp_Reset"
        Name = "Citrix Workspace App zuruecksetzen"
        Action = "WorkspaceApp_Reset"
        Category = "Citrix"
        ConfirmMessage = "Citrix Workspace App Einstellungen werden geloescht. Fortfahren?"
        Description = "Setzt Citrix Workspace-App-Einstellungen zurueck."
        Params = @{}
    },
    @{
        Id = "MS_Edge_Reset"
        Name = "Microsoft Edge Browser zuruecksetzen"
        Action = "MS_Edge_Reset"
        Category = "Browser"
        ConfirmMessage = "Edge-Einstellungen und lokale Browserdaten werden geloescht. Fortfahren?"
        Description = "Bereinigt Edge Cache, Profile und Konfigurationen."
        Params = @{}
    },
    @{
        Id = "Google_Chrome_Reset"
        Name = "Google Chrome Browser zuruecksetzen"
        Action = "Google_Chrome_Reset"
        Category = "Browser"
        ConfirmMessage = "Chrome-Einstellungen und lokale Browserdaten werden geloescht. Fortfahren?"
        Description = "Bereinigt Chrome Cache, Profile und Konfigurationen."
        Params = @{}
    },
    @{
        Id = "AdobeReader_Reset"
        Name = "Adobe Reader zuruecksetzen"
        Action = "AdobeReader_Reset"
        Category = "PDF"
        ConfirmMessage = "Adobe Reader Einstellungen werden geloescht. Fortfahren?"
        Description = "Setzt Adobe Reader Profil und Konfigurationen zurueck."
        Params = @{}
    },
    @{
        Id = "Outlook_Reset"
        Name = "Microsoft Outlook zuruecksetzen"
        Action = "Outlook_Reset"
        Category = "Office"
        ConfirmMessage = ""
        Description = "Bereinigt Outlook Cache, Registry-Reste und Suchindex (Benutzerkontext)."
        Params = @{}
    }
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Wartung ausfuehren (Benutzer)"
$form.StartPosition = "CenterScreen"
$form.AutoSize = $true
$form.AutoSizeMode = "GrowAndShrink"
$form.MaximizeBox = $false

$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Dock = 'Fill'
$flow.FlowDirection = 'TopDown'
$flow.WrapContents = $false
$flow.AutoSize = $true

$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $true
$label.Text = "Aktion(en) auswaehlen:"
$flow.Controls.Add($label)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000

$checkboxMap = @{}

foreach ($action in $actions) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $action.Name
    $cb.AutoSize = $true
    $cb.Tag = $action

    if ($action.Description) {
        $toolTip.SetToolTip($cb, $action.Description)
    }

    $flow.Controls.Add($cb)
    $checkboxMap[$action.Id] = $cb
}

$button = New-Object System.Windows.Forms.Button
$button.Text = "Ausfuehren"
$button.AutoSize = $true

$logButton = New-Object System.Windows.Forms.Button
$logButton.Text = "Log-Ordner oeffnen"
$logButton.AutoSize = $true

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.FlowDirection = 'LeftToRight'
$buttonPanel.WrapContents = $false
$buttonPanel.AutoSize = $true
$buttonPanel.Margin = '0,10,0,0'

$buttonPanel.Controls.Add($button)
$buttonPanel.Controls.Add($logButton)
$flow.Controls.Add($buttonPanel)

function Invoke-SelectedActions {
    $anySelected = $false

    foreach ($action in $actions) {
        $cb = $checkboxMap[$action.Id]
        if (-not $cb.Checked) { continue }

        $anySelected = $true
        Write-UiLog "Aktion gewaehlt" "INFO" $action.Action

        if ($action.ConfirmMessage) {
            $res = [System.Windows.Forms.MessageBox]::Show(
                $action.ConfirmMessage,
                $action.Name,
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-UiLog "Aktion abgebrochen" "WARN" $action.Action
                continue
            }
        }

        try {
            $result = Invoke-Action -Name $action.Action -Params $action.Params -Mode "Interactive" -ToolId $toolId -Trigger "UserUI"
            if ($result.ExitCode -ne 0 -or $result.Error) {
                Write-UiLog "Aktion fehlgeschlagen (ExitCode=$($result.ExitCode))" "WARN" $action.Action
            } else {
                Write-UiLog "Aktion abgeschlossen" "INFO" $action.Action
            }
        } catch {
            Write-UiLog $_.Exception.Message "ERROR" $action.Action
        }
    }

    return $anySelected
}

$button.Add_Click({
    if (Invoke-SelectedActions) {
        Write-UiLog "Aktionen abgeschlossen" "INFO"
        [System.Windows.Forms.MessageBox]::Show("Aktionen abgeschlossen.", "Fertig", 'OK', 'Information') | Out-Null
        foreach ($cb in $checkboxMap.Values) { $cb.Checked = $false }
    } else {
        Write-UiLog "Keine Aktion ausgewaehlt" "WARN"
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Aktion auswaehlen.", "Hinweis", 'OK', 'Warning') | Out-Null
    }

    if ($customer.logging -and $customer.logging.adminLogRoot) {
        $adminRoot = [string]$customer.logging.adminLogRoot
        if ($adminRoot -and (Test-Path $adminRoot)) {
            try {
                $userRoot = Join-Path $adminRoot $env:USERNAME
                $targetRoot = Join-Path $userRoot $toolId
                $targetSession = Join-Path $targetRoot $sessionId
                New-Item -ItemType Directory -Path $targetSession -Force | Out-Null

                if (Test-Path $logRoot) {
                    Copy-Item -Path $logRoot -Destination $targetSession -Recurse -Force -ErrorAction Stop
                    Write-UiLog "Log-Kopie nach Admin-Root erstellt: $targetSession" "INFO"
                }
            } catch {
                Write-UiLog "Admin-Log-Kopie fehlgeschlagen: $($_.Exception.Message)" "WARN"
            }
        } else {
            Write-UiLog "AdminLogRoot nicht erreichbar: $adminRoot" "WARN"
        }
    }
})

$logButton.Add_Click({
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    Start-Process explorer.exe "`"$logRoot`""
})

$form.Controls.Add($flow)
[void]$form.ShowDialog()
Write-UiLog "GUI geschlossen" "INFO"
