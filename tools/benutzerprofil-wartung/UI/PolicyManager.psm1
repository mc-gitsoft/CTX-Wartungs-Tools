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
        offline = [pscustomobject]@{
            every = [pscustomobject]@{ enabled = $false; actions = @() }
            once = @()
        }
    }
}

function Get-DefaultCustomer {
    return [pscustomobject]@{
        customer = [pscustomobject]@{ name = "" }
        paths = [pscustomobject]@{ repoRoot = "" }
        fslogix = [pscustomobject]@{ enabled = $false }
        branding = [pscustomobject]@{ windowTitle = "Wartung Admin"; supportText = "" }
        logging = [pscustomobject]@{ relativeLogPath = "logs"; adminLogRoot = "" }
        flags = [pscustomobject]@{ allowOffline = $true; allowLogoffRunner = $true }
    }
}

function Normalize-Actions {
    param([object]$Value)

    if (-not $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return @($Value)
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

function Get-TargetLines {
    param([object]$Value)

    if (-not $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return @($Value)
}

function Get-Lines {
    param([string]$Text)

    if (-not $Text) { return @() }
    return $Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Get-ActionNames {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot
    )

    $actionsPath = Join-Path $ToolRoot "Actions"
    if (-not (Test-Path $actionsPath)) { return @() }
    return Get-ChildItem -Path $actionsPath -Filter "*.ps1" | Sort-Object Name | ForEach-Object {
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    }
}

function Get-RecentCampaignIdsPath {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot
    )
    return (Join-Path $ToolRoot "recent_campaignIds.json")
}

function Load-RecentCampaignIds {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot
    )
    $path = Get-RecentCampaignIdsPath -ToolRoot $ToolRoot
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
        return @()
    }
}

function Save-RecentCampaignIds {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot,
        [string]$CampaignId
    )

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

    $path = Get-RecentCampaignIdsPath -ToolRoot $ToolRoot
    $json = @($recentList) | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Save-PolicyToFile {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot,
        [Parameter(Mandatory)]
        [object]$Policy
    )

    $path = Join-Path $ToolRoot "policy.json"
    $json = $Policy | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Load-PolicyFromFile {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot
    )

    $path = Join-Path $ToolRoot "policy.json"
    if (-not (Test-Path $path)) {
        return Get-DefaultPolicy
    }

    try {
        return (Get-Content $path -Raw | ConvertFrom-Json)
    }
    catch {
        Write-Warning ("Failed to parse policy.json: {0}" -f $_.Exception.Message)
        return Get-DefaultPolicy
    }
}

function Save-CustomerToFile {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot,
        [Parameter(Mandatory)]
        [object]$Customer
    )

    $path = Join-Path $ToolRoot "customer.json"
    $json = $Customer | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Load-CustomerFromFile {
    param(
        [Parameter(Mandatory)]
        [string]$ToolRoot
    )

    $path = Join-Path $ToolRoot "customer.json"
    if (-not (Test-Path $path)) {
        return Get-DefaultCustomer
    }

    try {
        return (Get-Content $path -Raw | ConvertFrom-Json)
    }
    catch {
        Write-Warning ("Failed to parse customer.json: {0}" -f $_.Exception.Message)
        return Get-DefaultCustomer
    }
}

Export-ModuleMember -Function Get-DefaultPolicy, Get-DefaultCustomer, `
    Normalize-Actions, Normalize-Once, Normalize-Every, `
    Get-TargetLines, Get-Lines, Get-ActionNames, `
    Get-RecentCampaignIdsPath, Load-RecentCampaignIds, Save-RecentCampaignIds, `
    Save-PolicyToFile, Load-PolicyFromFile, `
    Save-CustomerToFile, Load-CustomerFromFile
