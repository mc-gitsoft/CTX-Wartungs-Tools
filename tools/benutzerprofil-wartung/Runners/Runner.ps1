param(
    [ValidateSet("Logon","Logoff")]
    [string]$Trigger,
    [switch]$PreviewOnly
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot "shared\WartungsTools.SDK.psm1") -Force

$toolManifestPath = Join-Path $toolRoot "tool.json"
$toolManifest = Get-Content $toolManifestPath -Raw | ConvertFrom-Json
$toolId = $toolManifest.toolId

$policy = Get-PolicyConfig
$section = if ($Trigger -eq "Logon") { $policy.logon } else { $policy.logoff }

$stateRoot = Join-Path $env:LOCALAPPDATA "CTX-Wartungs-Tools\State\$toolId"

Write-Log -Level "INFO" -Message ("Runner start: Trigger={0}; ToolId={1}; ToolRoot={2}; LOCALAPPDATA={3}" -f $Trigger, $toolId, $toolRoot, $env:LOCALAPPDATA) -ToolId $toolId -Trigger $Trigger

function ConvertTo-Hashtable {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject
    )

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [hashtable]) { return $InputObject }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $ht = @{}
        foreach ($key in $InputObject.Keys) {
            $ht[$key] = ConvertTo-Hashtable -InputObject $InputObject[$key]
        }
        return $ht
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $list = @()
        foreach ($item in $InputObject) {
            $list += ConvertTo-Hashtable -InputObject $item
        }
        return $list
    }

    if ($InputObject -is [pscustomobject]) {
        $ht = @{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $ht[$prop.Name] = ConvertTo-Hashtable -InputObject $prop.Value
        }
        return $ht
    }

    return $InputObject
}

function Test-TargetMatch {
    param(
        [object]$Targets
    )
    if (-not $Targets) { return $true }

    $normalized = ConvertTo-Hashtable -InputObject $Targets
    if (-not $normalized) { return $true }

    $user = if ($env:USERDOMAIN) { "$env:USERDOMAIN\\$env:USERNAME" } else { $env:USERNAME }

    $users = $normalized.users
    if ($users) {
        if ($users -isnot [System.Collections.IEnumerable] -or $users -is [string]) {
            $users = @($users)
        }

        if (($users -notcontains $user) -and ($users -notcontains $env:USERNAME)) {
            return $false
        }
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

        Write-Log -Level "INFO" -Message ("Evaluate {0}: campaignId={1}; validUntil={2}" -f $block.Name, $campaignId, $validUntil) -ToolId $toolId -Trigger $Trigger

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

        $targetsMatch = Test-TargetMatch -Targets $item.targets
        Write-Log -Level "INFO" -Message ("Targets match: {0}" -f $targetsMatch) -ToolId $toolId -Trigger $Trigger
        if (-not $targetsMatch) {
            Write-Log -Level "INFO" -Message "Skip: targets do not match." -ToolId $toolId -Trigger $Trigger
            continue
        }

        $stateFile = $null
        if ($block.Name -eq "once") {
            if (-not $campaignId) {
                Write-Log -Level "WARN" -Message "Missing campaignId for once block. Running without state." -ToolId $toolId -Trigger $Trigger
            } else {
                $stateFile = Join-Path $stateRoot ("{0}_once_{1}.json" -f $Trigger.ToLowerInvariant(), $campaignId)
                if (Test-Path $stateFile) {
                    Write-Log -Level "INFO" -Message ("Skip: state already exists: {0}" -f $stateFile) -ToolId $toolId -Trigger $Trigger
                    continue
                }
            }
        }

        $hadErrors = $false
        $actionResults = @()
        foreach ($action in $item.actions) {
            try {
                if ($PreviewOnly) {
                    Write-Log -Level "INFO" -Message ("PREVIEW: WOULD RUN: {0}" -f $action.name) -ToolId $toolId -Trigger $Trigger -Action $action.name
                    $actionResults += [pscustomobject]@{
                        Name = $action.name
                        ExitCode = 0
                        Error = $null
                    }
                } else {
                    Write-Log -Level "INFO" -Message ("Execute action: {0}" -f $action.name) -ToolId $toolId -Trigger $Trigger -Action $action.name
                    $result = Invoke-Action -Name $action.name -Params $action.params -Mode "Silent" -ForceSilent -ToolId $toolId -Trigger $Trigger
                    if ($result.ExitCode -ne 0 -or $result.Error) {
                        $hadErrors = $true
                        Write-Log -Level "INFO" -Message ("Action finished with errors: {0}; ExitCode={1}" -f $action.name, $result.ExitCode) -ToolId $toolId -Trigger $Trigger -Action $action.name
                    } else {
                        Write-Log -Level "INFO" -Message ("Action finished OK: {0}" -f $action.name) -ToolId $toolId -Trigger $Trigger -Action $action.name
                    }
                    $actionResults += $result
                }
            } catch {
                $hadErrors = $true
                if ($PreviewOnly) {
                    Write-Log -Level "INFO" -Message ("PREVIEW: SKIP (error): {0}" -f $action.name) -ToolId $toolId -Trigger $Trigger -Action $action.name
                    $actionResults += [pscustomobject]@{
                        Name = $action.name
                        ExitCode = 0
                        Error = $null
                    }
                } else {
                    Write-Log -Level "ERROR" -Message $_.Exception.Message -ToolId $toolId -Trigger $Trigger -Action $action.name
                    $actionResults += [pscustomobject]@{
                        Name = $action.name
                        ExitCode = 1
                        Error = $_.Exception.Message
                    }
                }
            }
        }

        if ($block.Name -eq "once" -and $stateFile -and -not $hadErrors -and -not $PreviewOnly) {
            try {
                Write-Log -Level "INFO" -Message ("Write state: Root={0}; File={1}" -f $stateRoot, $stateFile) -ToolId $toolId -Trigger $Trigger
                New-Item -ItemType Directory -Path $stateRoot -Force | Out-Null
                $stateRootExists = Test-Path $stateRoot
                Write-Log -Level "INFO" -Message ("State root exists: {0}" -f $stateRootExists) -ToolId $toolId -Trigger $Trigger
                $state = [pscustomobject]@{
                    doneAt = (Get-Date).ToString("s")
                    campaignId = $campaignId
                    trigger = $Trigger
                    actions = $actionResults
                }
                $json = $state | ConvertTo-Json -Depth 6
                Set-Content -Path $stateFile -Value $json -Encoding UTF8
                if (Test-Path $stateFile) {
                    Write-Log -Level "INFO" -Message ("State write OK: {0}" -f $stateFile) -ToolId $toolId -Trigger $Trigger
                } else {
                    Write-Log -Level "INFO" -Message ("State write FAIL (missing after write): {0}" -f $stateFile) -ToolId $toolId -Trigger $Trigger
                }
            } catch {
                Write-Log -Level "ERROR" -Message ("Failed to write state: {0} | {1}" -f $stateFile, $_.Exception.Message) -ToolId $toolId -Trigger $Trigger
            }
        } elseif ($block.Name -eq "once" -and $stateFile -and $hadErrors -and -not $PreviewOnly) {
            Write-Log -Level "INFO" -Message ("Skip state write due to errors: {0}" -f $stateFile) -ToolId $toolId -Trigger $Trigger
        }
    }
}
