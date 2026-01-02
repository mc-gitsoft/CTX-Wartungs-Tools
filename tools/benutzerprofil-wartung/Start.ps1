param(
    [ValidateSet("User","Admin")]
    [string]$UiMode = "User"
)

$toolRoot = Split-Path -Parent $PSCommandPath
$modulePath = Join-Path $toolRoot "shared\WartungsTools.SDK.psm1"
Import-Module $modulePath -Force

switch ($UiMode) {
    "Admin" { & (Join-Path $toolRoot "UI\Wartung_Admin.ps1") }
    default { & (Join-Path $toolRoot "UI\Wartung_User.ps1") }
}
