$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$source = Join-Path $repoRoot "shared\SDK\WartungsTools.SDK.psm1"
$toolsRoot = Join-Path $repoRoot "tools"

if (-not (Test-Path $source)) {
    throw "SDK source not found: $source"
}

Get-ChildItem -Path $toolsRoot -Directory | ForEach-Object {
    $destDir = Join-Path $_.FullName "shared"
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    Copy-Item -Path $source -Destination (Join-Path $destDir "WartungsTools.SDK.psm1") -Force
}

Write-Output "SDK synced to tools/*/shared"
