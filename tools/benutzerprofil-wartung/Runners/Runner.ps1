param(
    [ValidateSet("Logon","Logoff")]
    [string]$Trigger
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot "shared\WartungsTools.SDK.psm1") -Force

$toolManifestPath = Join-Path $toolRoot "tool.json"
$toolManifest = Get-Content $toolManifestPath -Raw | ConvertFrom-Json
$toolId = $toolManifest.toolId

$policy = Get-PolicyConfig
$section = if ($Trigger -eq "Logon") { $policy.logon } else { $policy.logoff }

$stateRoot = Join-Path $env:LOCALAPPDATA "CTX-Wartungs-Tools\State\$toolId"

function Test-TargetMatch {
    param(
        [hashtable]$Targets
    )
    if (-not $Targets) { return $true }

    $user = if ($env:USERDOMAIN) { "$env:USERDOMAIN\\$env:USERNAME" } else { $env:USERNAME }

    if ($Targets.users -and ($Targets.users -notcontains $user) -and ($Targets.users -notcontains $env:USERNAME)) {
        return $false
    }

    return $true
}

$blocks = @(
    @{ Name = "every"; Items = $section.every },
    @{ Name = "once"; Items = $section.once }
)

foreach ($block in $blocks) {
    foreach ($item in $block.Items) {
        $campaignId = $item.campaignId
        $validUntil = $item.validUntil

        if ($validUntil) {
            try {
                $until = [datetime]::Parse($validUntil)
                if ($until -lt (Get-Date)) {
                    Write-Log -Level "WARN" -Message "Campaign expired: $campaignId" -ToolId $toolId -Trigger $Trigger
                    continue
                }
            } catch {
                Write-Log -Level "WARN" -Message "Invalid validUntil date: $validUntil" -ToolId $toolId -Trigger $Trigger
            }
        }

        if (-not (Test-TargetMatch -Targets $item.targets)) {
            continue
        }

        $stateFile = $null
        if ($block.Name -eq "once") {
            if (-not $campaignId) {
                Write-Log -Level "WARN" -Message "Missing campaignId for once block. Running without state." -ToolId $toolId -Trigger $Trigger
            } else {
                $stateFile = Join-Path $stateRoot "$campaignId.state"
                if (Test-Path $stateFile) {
                    continue
                }
            }
        }

        $hadErrors = $false
        foreach ($action in $item.actions) {
            try {
                $result = Invoke-Action -Name $action.name -Params $action.params -Mode "Silent" -ForceSilent -ToolId $toolId -Trigger $Trigger
                if ($result.ExitCode -ne 0 -or $result.Error) {
                    $hadErrors = $true
                }
            } catch {
                $hadErrors = $true
                Write-Log -Level "ERROR" -Message $_.Exception.Message -ToolId $toolId -Trigger $Trigger -Action $action.name
            }
        }

        if ($block.Name -eq "once" -and $stateFile -and -not $hadErrors) {
            New-Item -ItemType Directory -Path $stateRoot -Force | Out-Null
            Set-Content -Path $stateFile -Value (Get-Date).ToString("s") -Encoding UTF8
        }
    }
}
