param(
    [ValidateSet("Logon","Logoff","Offline")]
    [string]$Trigger,
    [switch]$PreviewOnly,
    [string]$TargetUser,
    [string]$VhdPath
)

$toolRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
Import-Module (Join-Path $toolRoot "shared\WartungsTools.SDK.psm1") -Force

$toolManifestPath = Join-Path $toolRoot "tool.json"
try {
    $toolManifest = Get-Content $toolManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
    $toolId = $toolManifest.toolId
}
catch {
    Write-Warning ("Failed to load tool.json: {0}" -f $_.Exception.Message)
    exit 1
}

$policy = Get-PolicyConfig
$section = switch ($Trigger) {
    "Logon"  { $policy.logon }
    "Logoff" { $policy.logoff }
    "Offline" { $policy.offline }
}
if (-not $section) {
    Write-Log -Level "WARN" -Message "No policy section found for trigger: $Trigger" -ToolId $toolId -Trigger $Trigger
    exit 0
}

$stateRoot = Join-Path $env:LOCALAPPDATA "CTX-Wartungs-Tools\State\$toolId"

function Resolve-OfflineProfileRoot {
    param(
        [Parameter(Mandatory)]
        [string]$User
    )

    $sanitized = ($User -replace "\\","_")
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalid) {
        $sanitized = $sanitized.Replace($ch, "_")
    }
    return ("C:\\_OfflineProfiles\\{0}" -f $sanitized)
}

if ($Trigger -eq "Offline" -and -not $TargetUser -and -not $VhdPath) {
    throw "Offline trigger requires -VhdPath or -TargetUser."
}

# Extract username from VHD path if TargetUser not provided
if ($Trigger -eq "Offline" -and -not $TargetUser -and $VhdPath) {
    $vhdName = [System.IO.Path]::GetFileNameWithoutExtension($VhdPath)
    if ($vhdName -match '^Profile_(.+)$') {
        $TargetUser = $Matches[1]
    } else {
        $parentDir = Split-Path (Split-Path $VhdPath -Parent) -Leaf
        if ($parentDir -match '_(.+)$') {
            $TargetUser = $Matches[1]
        } else {
            $TargetUser = $parentDir
        }
    }
    Write-Log -Level "INFO" -Message ("TargetUser derived from VHD path: {0}" -f $TargetUser) -ToolId $toolId -Trigger $Trigger
}

$customerConfig = Get-CustomerConfig

$contextProfileRoot = $env:USERPROFILE
$savedAppData = $env:APPDATA
$savedLocalAppData = $env:LOCALAPPDATA
$offlineMode = $false

if ($Trigger -eq "Offline") {
    if (-not $customerConfig.flags.allowOffline) {
        Write-Log -Level "WARN" -Message "Offline mode disabled in customer.json (flags.allowOffline = false)" -ToolId $toolId -Trigger $Trigger
        exit 0
    }

    $mountedVhd = $null

    if ($VhdPath) {
        # Mount directly specified VHD(X) container
        if (-not (Test-Path $VhdPath)) {
            Write-Log -Level "ERROR" -Message ("VHD not found: {0}" -f $VhdPath) -ToolId $toolId -Trigger $Trigger
            exit 1
        }
        try {
            $mountPath = Mount-FSLogixVHD -VhdPath $VhdPath
            $contextProfileRoot = Join-Path $mountPath "Profile"
            $mountedVhd = $VhdPath
            Write-Log -Level "INFO" -Message ("VHD mounted: {0} -> {1}" -f $VhdPath, $contextProfileRoot) -ToolId $toolId -Trigger $Trigger
        } catch {
            Write-Log -Level "ERROR" -Message ("VHD mount failed: {0}" -f $_.Exception.Message) -ToolId $toolId -Trigger $Trigger
            exit 1
        }
    } else {
        # Fallback to static offline profile directory
        $contextProfileRoot = Resolve-OfflineProfileRoot -User $TargetUser
        if (-not (Test-Path $contextProfileRoot)) {
            Write-Log -Level "ERROR" -Message ("Offline profile root not found: {0}" -f $contextProfileRoot) -ToolId $toolId -Trigger $Trigger
            exit 1
        }
    }

    $offlineMode = $true
    $env:APPDATA = Join-Path $contextProfileRoot "AppData\Roaming"
    $env:LOCALAPPDATA = Join-Path $contextProfileRoot "AppData\Local"
    Write-Log -Level "INFO" -Message ("Offline mode: TargetUser={0}; ProfileRoot={1}; APPDATA={2}; LOCALAPPDATA={3}" -f $TargetUser, $contextProfileRoot, $env:APPDATA, $env:LOCALAPPDATA) -ToolId $toolId -Trigger $Trigger
}

$Context = @{
    Trigger = $Trigger
    TargetUser = $TargetUser
    ProfileRoot = $contextProfileRoot
    OfflineMode = $offlineMode
}

$env:CTX_TRIGGER = $Context.Trigger
$env:CTX_TARGET_USER = $Context.TargetUser
$env:CTX_PROFILE_ROOT = $Context.ProfileRoot
$env:CTX_OFFLINE = if ($offlineMode) { "1" } else { "0" }

Write-Log -Level "INFO" -Message ("Runner start: Trigger={0}; ToolId={1}; ToolRoot={2}; LOCALAPPDATA={3}; TargetUser={4}" -f $Trigger, $toolId, $toolRoot, $env:LOCALAPPDATA, $TargetUser) -ToolId $toolId -Trigger $Trigger
Write-Log -Level "INFO" -Message ("Context: CTX_TRIGGER={0}; CTX_TARGET_USER={1}; CTX_PROFILE_ROOT={2}" -f $env:CTX_TRIGGER, $env:CTX_TARGET_USER, $env:CTX_PROFILE_ROOT) -ToolId $toolId -Trigger $Trigger

function Test-TargetMatch {
    param(
        [object]$Targets
    )
    if (-not $Targets) { return $true }

    $normalized = ConvertTo-Hashtable -InputObject $Targets
    if (-not $normalized) { return $true }

    $user = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME }

    # Check user targets
    $users = $normalized.users
    if ($users) {
        if ($users -isnot [System.Collections.IEnumerable] -or $users -is [string]) {
            $users = @($users)
        }

        if (($users -notcontains $user) -and ($users -notcontains $env:USERNAME)) {
            return $false
        }
    }

    # Check group targets
    $groups = $normalized.groups
    if ($groups) {
        if ($groups -isnot [System.Collections.IEnumerable] -or $groups -is [string]) {
            $groups = @($groups)
        }

        $userGroups = @()
        try {
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($identity)
            $userGroups = $identity.Groups | ForEach-Object {
                try { $_.Translate([Security.Principal.NTAccount]).Value } catch { $null }
            } | Where-Object { $_ }
        }
        catch {
            Write-Log -Level "WARN" -Message ("Could not resolve group membership: {0}" -f $_.Exception.Message) -ToolId $toolId -Trigger $Trigger
            return $true
        }

        $matchFound = $false
        foreach ($targetGroup in $groups) {
            if ($userGroups -contains $targetGroup) {
                $matchFound = $true
                break
            }
            # Also check without domain prefix
            $shortGroups = $userGroups | ForEach-Object { ($_ -split '\\')[-1] }
            $shortTarget = ($targetGroup -split '\\')[-1]
            if ($shortGroups -contains $shortTarget) {
                $matchFound = $true
                break
            }
        }

        if (-not $matchFound) {
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
                Write-Log -Level "ERROR" -Message "Missing campaignId for once block. Skipping to prevent untracked execution." -ToolId $toolId -Trigger $Trigger
                continue
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

                $state = [pscustomobject]@{
                    doneAt = (Get-Date).ToString("s")
                    campaignId = $campaignId
                    trigger = $Trigger
                    actions = $actionResults
                }
                $json = $state | ConvertTo-Json -Depth 6

                # Atomic write: write to temp file first, then rename
                $tempFile = $stateFile + ".tmp"
                Set-Content -Path $tempFile -Value $json -Encoding UTF8 -ErrorAction Stop

                if (Test-Path $stateFile) {
                    Remove-Item -Path $stateFile -Force -ErrorAction SilentlyContinue
                }
                Move-Item -Path $tempFile -Destination $stateFile -Force -ErrorAction Stop

                Write-Log -Level "INFO" -Message ("State write OK: {0}" -f $stateFile) -ToolId $toolId -Trigger $Trigger
            } catch {
                Write-Log -Level "ERROR" -Message ("Failed to write state: {0} | {1}" -f $stateFile, $_.Exception.Message) -ToolId $toolId -Trigger $Trigger
                # Clean up temp file if it exists
                if (Test-Path ($stateFile + ".tmp")) {
                    Remove-Item -Path ($stateFile + ".tmp") -Force -ErrorAction SilentlyContinue
                }
            }
        } elseif ($block.Name -eq "once" -and $stateFile -and $hadErrors -and -not $PreviewOnly) {
            Write-Log -Level "WARN" -Message ("Skip state write due to errors: {0}" -f $stateFile) -ToolId $toolId -Trigger $Trigger
        }
    }
}

# Restore environment variables and dismount VHD after offline execution
if ($offlineMode) {
    $env:APPDATA = $savedAppData
    $env:LOCALAPPDATA = $savedLocalAppData

    if ($mountedVhd) {
        [gc]::Collect()
        Start-Sleep -Milliseconds 500
        $dismountOk = Dismount-FSLogixVHD -VhdPath $mountedVhd
        if ($dismountOk) {
            Write-Log -Level "INFO" -Message ("FSLogix VHD dismounted: {0}" -f $mountedVhd) -ToolId $toolId -Trigger $Trigger
        } else {
            Write-Log -Level "WARN" -Message ("FSLogix VHD dismount failed: {0}" -f $mountedVhd) -ToolId $toolId -Trigger $Trigger
        }
    }

    Write-Log -Level "INFO" -Message "Offline mode: environment restored." -ToolId $toolId -Trigger $Trigger
}
