function Get-ToolRoot {
    param(
        [string]$ScriptPath
    )

    $path = if ($ScriptPath) { $ScriptPath } elseif ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.PSCommandPath) { $MyInvocation.PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if (-not $path) {
        throw "Unable to determine script path for tool root detection."
    }

    $dir = Split-Path -Parent $path
    while ($dir -and -not (Test-Path (Join-Path $dir "tool.json"))) {
        $parent = Split-Path -Parent $dir
        if ($parent -eq $dir) { $dir = $null; break }
        $dir = $parent
    }

    if (-not $dir) {
        throw "Tool root not found. Ensure tool.json exists in the tool root."
    }

    return $dir
}

function Get-CustomerConfig {
    $toolRoot = Get-ToolRoot
    $path = Join-Path $toolRoot "customer.json"

    if (-not (Test-Path $path)) {
        throw "customer.json not found. Copy customer.json.example to customer.json in the tool root."
    }

    return (Get-Content $path -Raw | ConvertFrom-Json)
}

function Get-PolicyConfig {
    $toolRoot = Get-ToolRoot
    $path = Join-Path $toolRoot "policy.json"

    $default = [pscustomobject]@{
        logon  = [pscustomobject]@{ every = @(); once = @() }
        logoff = [pscustomobject]@{ every = @(); once = @() }
    }

    if (-not (Test-Path $path)) {
        return $default
    }

    try {
        return (Get-Content $path -Raw | ConvertFrom-Json)
    }
    catch {
        Write-Warning ("Failed to parse policy.json: {0}" -f $_.Exception.Message)
        return $default
    }
}

function Write-Log {
    param(
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO",
        [string]$Message,
        [string]$ToolId,
        [string]$Trigger,
        [string]$Action,
        [string]$LogRoot
    )

    $toolRoot = Get-ToolRoot
    $relativeLogPath = "logs"

    if (-not $LogRoot) {
        try {
            $customer = Get-CustomerConfig
            if ($customer.logging.relativeLogPath) {
                $relativeLogPath = $customer.logging.relativeLogPath
            }
        } catch {
            $relativeLogPath = "logs"
        }
    }

    $basePath = if ($LogRoot) { $LogRoot } else { Join-Path $toolRoot $relativeLogPath }
    New-Item -ItemType Directory -Path $basePath -Force | Out-Null

    $dateStamp = (Get-Date).ToString("yyyyMMdd")
    $filePath = Join-Path $basePath ("$ToolId-$dateStamp.log")

    $timestamp = (Get-Date).ToString("s")
    $context = @()
    if ($ToolId) { $context += "tool=$ToolId" }
    if ($Trigger) { $context += "trigger=$Trigger" }
    if ($Action) { $context += "action=$Action" }

    $line = "[$timestamp][$Level] $Message"
    if ($context.Count -gt 0) {
        $line = $line + " | " + ($context -join ";")
    }

    Add-Content -Path $filePath -Value $line -Encoding UTF8
}

function ConvertTo-Hashtable {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject
    )

    if ($null -eq $InputObject) { return @{} }

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

function Stop-SessionProcesses {
    <#
    .SYNOPSIS
        Beendet Prozesse anhand ihrer Namen in der aktuellen User-Session (RDS/Citrix-safe).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ProcessNames,

        [int]$Retries = 5,
        [int]$DelayMs = 400,

        [switch]$UseTaskkillFallback
    )

    # Skip process termination in offline mode (no active user session)
    if ($env:CTX_OFFLINE -eq "1") {
        return 0
    }

    try {
        $sessionId = (Get-Process -Id $PID).SessionId
    }
    catch {
        $sessionId = $null
    }

    $normalized = $ProcessNames | ForEach-Object { $_.ToLower() } | Select-Object -Unique
    $killed = 0

    for ($i = 1; $i -le $Retries; $i++) {
        try {
            if ($null -ne $sessionId) {
                $candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                    ($normalized -contains $_.Name.ToLower()) -and ($_.SessionId -eq $sessionId)
                }
            }
            else {
                $candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                    $normalized -contains $_.Name.ToLower()
                }
            }

            if (-not $candidates) {
                if ($i -eq 1) { break }
                continue
            }

            foreach ($p in $candidates) {
                try {
                    Stop-Process -Id $p.Id -Force -ErrorAction Stop
                    $killed++
                }
                catch {
                    # Ignore per-process errors during retry loop
                }
            }

            if ($UseTaskkillFallback) {
                $still = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                    ($normalized -contains $_.Name.ToLower()) -and
                    (($null -eq $sessionId) -or ($_.SessionId -eq $sessionId))
                }
                foreach ($p in $still) {
                    try {
                        cmd.exe /c "taskkill /F /T /PID $($p.Id)" | Out-Null
                        $killed++
                    }
                    catch { }
                }
            }
        }
        catch { }

        Start-Sleep -Milliseconds $DelayMs
    }

    return $killed
}

function Remove-PathSafe {
    <#
    .SYNOPSIS
        Robuste Ordner-/Datei-Löschung mit mehrstufigem Fallback (cmd rd, robocopy mirror, Remove-Item).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    $origLoc = Get-Location
    try { Set-Location -LiteralPath $env:SystemRoot } catch { }

    try {
        # 1) cmd rd
        try { & cmd.exe /d /c "rd /s /q ""$Path"" >nul 2>nul" | Out-Null } catch { }
        if (-not (Test-Path -LiteralPath $Path)) { return $true }

        # 2) robocopy mirror
        $empty = Join-Path $env:TEMP ("_empty_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $empty -Force | Out-Null
        try {
            & robocopy.exe $empty $Path /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP > $null 2> $null
        }
        catch { }
        try { Remove-Item -LiteralPath $empty -Recurse -Force -ErrorAction SilentlyContinue } catch { }

        # 3) Final cleanup
        try { & cmd.exe /d /c "rd /s /q ""$Path"" >nul 2>nul" | Out-Null } catch { }
        if (Test-Path -LiteralPath $Path) {
            try { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        }

        return (-not (Test-Path -LiteralPath $Path))
    }
    catch {
        return (-not (Test-Path -LiteralPath $Path))
    }
    finally {
        try { Set-Location -LiteralPath $origLoc.Path } catch { }
    }
}

function Clear-RegistryPath {
    <#
    .SYNOPSIS
        Entfernt einen HKCU-Registry-Schlüssel rekursiv und sicher.
        Im Offline-Modus wird NTUSER.DAT temporaer geladen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    # Offline mode: load NTUSER.DAT, remap path, operate, unload
    if ($env:CTX_OFFLINE -eq "1" -and $env:CTX_PROFILE_ROOT -and $Path -match '^HKCU:\\') {
        $ntUserDat = Join-Path $env:CTX_PROFILE_ROOT "NTUSER.DAT"
        if (-not (Test-Path $ntUserDat)) { return $true }

        $hiveKey = "HKU\CTX_OFFLINE_HIVE"
        $regPath = $Path -replace '^HKCU:\\', ''

        try {
            & reg load $hiveKey $ntUserDat 2>$null
            $offlinePath = "Registry::HKEY_USERS\CTX_OFFLINE_HIVE\$regPath"
            if (Test-Path $offlinePath) {
                Remove-Item -Path $offlinePath -Recurse -Force -ErrorAction Stop
            }
            return $true
        } catch {
            return $false
        } finally {
            [gc]::Collect()
            & reg unload $hiveKey 2>$null
        }
    }

    if (-not (Test-Path $Path)) {
        return $true
    }

    try {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Invoke-Action {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [object]$Params = @{},
        [ValidateSet("Silent","Interactive")]
        [string]$Mode = "Interactive",
        [switch]$ForceSilent,
        [string]$ToolId,
        [string]$Trigger
    )

    if ($ForceSilent) { $Mode = "Silent" }

    $Params = ConvertTo-Hashtable -InputObject $Params
    if ($Params -isnot [hashtable]) { $Params = @{} }

    $toolRoot = Get-ToolRoot
    $path = Join-Path $toolRoot "Actions\$Name.ps1"

    if (-not (Test-Path $path)) {
        $message = "Action script not found: $Name"
        Write-Log -Level "ERROR" -Message $message -ToolId $ToolId -Trigger $Trigger -Action $Name
        throw $message
    }

    $command = Get-Command -Name $path -ErrorAction Stop
    $paramNames = $command.Parameters.Keys
    $hasParams = $paramNames -contains "Params"
    $hasMode = $paramNames -contains "Mode"

    try {
        if ($hasParams) {
            if ($hasMode) {
                & $path -Params $Params -Mode $Mode
            } else {
                & $path -Params $Params
            }
        } else {
            $invokeParams = @{}
            if ($Params) { $invokeParams += $Params }
            if ($hasMode) { $invokeParams["Mode"] = $Mode }
            & $path @invokeParams
        }
        $exitCode = $LASTEXITCODE
        if ($null -eq $exitCode) { $exitCode = 0 }
        return [pscustomobject]@{
            Name = $Name
            ExitCode = $exitCode
            Error = $null
        }
    } catch {
        $err = $_.Exception.Message
        Write-Log -Level "ERROR" -Message $err -ToolId $ToolId -Trigger $Trigger -Action $Name
        return [pscustomobject]@{
            Name = $Name
            ExitCode = 1
            Error = $err
        }
    }
}

function Mount-FSLogixVHD {
    <#
    .SYNOPSIS
        Mountet ein FSLogix-Profil-VHD(X) und gibt den Mount-Pfad zurueck.
        Nutzt Mount-DiskImage (Storage-Modul, ueberall verfuegbar) statt Mount-VHD (Hyper-V).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$VhdPath,

        [switch]$ReadOnly
    )

    if (-not (Test-Path $VhdPath)) {
        throw "VHD not found: $VhdPath"
    }

    try {
        $access = if ($ReadOnly) { "ReadOnly" } else { "ReadWrite" }
        $diskImage = Mount-DiskImage -ImagePath $VhdPath -Access $access -PassThru -ErrorAction Stop
        $disk = $diskImage | Get-Disk -ErrorAction Stop
        $partitions = $disk | Get-Partition -ErrorAction Stop
        $volume = $partitions | Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }

        if ($volume) {
            return "$($volume.DriveLetter):\"
        }

        # No drive letter assigned — use access path
        $accessPath = ($partitions | Where-Object { $_.AccessPaths.Count -gt 0 } | Select-Object -First 1).AccessPaths[0]
        if ($accessPath) {
            return $accessPath
        }

        throw "No volume or access path found after mounting."
    } catch {
        try { Dismount-DiskImage -ImagePath $VhdPath -ErrorAction SilentlyContinue } catch { }
        throw "Failed to mount VHD: $($_.Exception.Message)"
    }
}

function Dismount-FSLogixVHD {
    <#
    .SYNOPSIS
        Dismountet ein FSLogix-Profil-VHD(X).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$VhdPath
    )

    try {
        Dismount-DiskImage -ImagePath $VhdPath -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

Export-ModuleMember -Function Get-ToolRoot, Get-CustomerConfig, Get-PolicyConfig, Write-Log, ConvertTo-Hashtable, Invoke-Action, Stop-SessionProcesses, Remove-PathSafe, Clear-RegistryPath, Mount-FSLogixVHD, Dismount-FSLogixVHD

