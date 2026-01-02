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

    if (-not (Test-Path $path)) {
        return [pscustomobject]@{
            logon  = [pscustomobject]@{ every = @(); once = @() }
            logoff = [pscustomobject]@{ every = @(); once = @() }
        }
    }

    return (Get-Content $path -Raw | ConvertFrom-Json)
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

Export-ModuleMember -Function Get-ToolRoot, Get-CustomerConfig, Get-PolicyConfig, Write-Log, ConvertTo-Hashtable, Invoke-Action

