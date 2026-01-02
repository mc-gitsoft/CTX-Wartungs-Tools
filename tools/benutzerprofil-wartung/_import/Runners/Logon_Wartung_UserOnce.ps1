[CmdletBinding()]
param()

# =========================
# 0) FIXE KONFIG-PFADE
# =========================
$ConfigPath = "\\kontor-n.local\NETLOGON\Citrix\Scripte\Wartungsscript\v2\AdminTool\Config\logon_once.json"

# Lokale Ablage (User darf schreiben)
$LocalRoot   = Join-Path $env:LOCALAPPDATA "Wartungsscript"
$CacheRoot   = Join-Path $LocalRoot "Cache"
$LogRoot     = Join-Path $LocalRoot "Logs\LogonOnce"
$ConfigCacheDir = Join-Path $LocalRoot "ConfigCache"
$LkgConfigPath  = Join-Path $ConfigCacheDir "logon_once.lastgood.json"

New-Item -Path $CacheRoot      -ItemType Directory -Force | Out-Null
New-Item -Path $LogRoot        -ItemType Directory -Force | Out-Null
New-Item -Path $ConfigCacheDir -ItemType Directory -Force | Out-Null

$SessionTs   = Get-Date -Format "yyyyMMdd_HHmmss"
$MasterLog   = Join-Path $LogRoot ("LogonOnce_{0}.log" -f $SessionTs)

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path $MasterLog -Value "$ts;$Level;$($env:COMPUTERNAME);$($env:USERNAME);$Message"
}

Write-Log "LogonOnce gestartet. ConfigPath=$ConfigPath"

# =========================
# 1) JSON LESEN (MIT TIMEOUT) + LKG FALLBACK
# =========================
function Get-ConfigJsonRaw {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 3
    )
    try {
        $job = Start-Job -ScriptBlock {
            param($p)
            if (Test-Path $p) { Get-Content -LiteralPath $p -Raw -ErrorAction Stop } else { $null }
        } -ArgumentList $Path

        if (Wait-Job $job -Timeout $TimeoutSeconds) {
            $content = Receive-Job $job -ErrorAction SilentlyContinue
            Remove-Job $job -Force -ErrorAction SilentlyContinue
            return $content
        } else {
            Stop-Job $job -Force -ErrorAction SilentlyContinue
            Remove-Job $job -Force -ErrorAction SilentlyContinue
            return $null
        }
    } catch {
        return $null
    }
}

function Load-Config {
    param(
        [Parameter(Mandatory)][string]$PrimaryPath,
        [Parameter(Mandatory)][string]$LastGoodPath,
        [int]$TimeoutSeconds = 3
    )

    # 1) Primär versuchen
    $raw = Get-ConfigJsonRaw -Path $PrimaryPath -TimeoutSeconds $TimeoutSeconds
    if ($raw) {
        try {
            $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
            # Wenn parse ok: als LKG ablegen (atomar-ish)
            try {
                $tmp = "$LastGoodPath.tmp"
                Set-Content -LiteralPath $tmp -Value $raw -Encoding UTF8 -Force
                Move-Item -LiteralPath $tmp -Destination $LastGoodPath -Force
                Write-Log "Config OK von Share geladen und als LastKnownGood gespeichert: $LastGoodPath"
            } catch {
                Write-Log "Konnte LastKnownGood nicht schreiben: $($_.Exception.Message)" "WARN"
            }
            return $cfg
        } catch {
            Write-Log "Config JSON parse Fehler (Share): $($_.Exception.Message)" "WARN"
            # Fallthrough -> LKG
        }
    } else {
        Write-Log "Config nicht erreichbar/leer (Share). Versuche LastKnownGood..." "WARN"
    }

    # 2) Fallback: LastKnownGood
    if (Test-Path -LiteralPath $LastGoodPath) {
        try {
            $lkgRaw = Get-Content -LiteralPath $LastGoodPath -Raw -ErrorAction Stop
            $cfg = $lkgRaw | ConvertFrom-Json -ErrorAction Stop
            Write-Log "Config aus LastKnownGood geladen: $LastGoodPath" "WARN"
            return $cfg
        } catch {
            Write-Log "LastKnownGood parse Fehler: $($_.Exception.Message)" "ERROR"
            return $null
        }
    } else {
        Write-Log "Keine LastKnownGood vorhanden: $LastGoodPath" "WARN"
        return $null
    }
}

# Timeout aus Default (wird ggf. erst nach Load-Config aus cfg überschrieben)
$defaultShareTimeout = 3
$cfg = Load-Config -PrimaryPath $ConfigPath -LastGoodPath $LkgConfigPath -TimeoutSeconds $defaultShareTimeout

if (-not $cfg) {
    Write-Log "Keine gültige Config (weder Share noch LKG). Wartung übersprungen." "WARN"
    exit 0
}

# =========================
# 2) SAFETY DEFAULTS + ENABLE / VALIDITY
# =========================
$maxRuntime = 600
$stopOnErr  = $false
$cleanupDays = 30

try { if ($cfg.safety.maxRuntimeSeconds) { $maxRuntime = [int]$cfg.safety.maxRuntimeSeconds } } catch {}
try { if ($cfg.safety.stopOnFirstError)  { $stopOnErr  = [bool]$cfg.safety.stopOnFirstError } } catch {}
try { if ($cfg.safety.cleanupDays)       { $cleanupDays = [int]$cfg.safety.cleanupDays } } catch {}

if (-not $cfg.enabled) {
    Write-Log "Config enabled=false -> Ende."
    exit 0
}

if ($cfg.validUntil) {
    try {
        $vu = [datetime]::ParseExact([string]$cfg.validUntil, "yyyy-MM-dd", $null)
        if ((Get-Date) -gt $vu.AddDays(1).AddSeconds(-1)) {
            Write-Log "validUntil überschritten ($($cfg.validUntil)) -> Ende." "WARN"
            exit 0
        }
    } catch {
        Write-Log "validUntil ungültig: $($cfg.validUntil) -> ignoriere." "WARN"
    }
}

if (-not $cfg.campaignId) {
    Write-Log "campaignId fehlt -> Ende." "WARN"
    exit 0
}

# =========================
# 3) CLEANUP (alte Campaign-Caches + Logs)
# =========================
function Cleanup-OldItems {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][int]$KeepDays,
        [string]$What = "Items"
    )

    try {
        if (-not (Test-Path -LiteralPath $Folder)) { return }
        $limit = (Get-Date).AddDays(-$KeepDays)

        Get-ChildItem -LiteralPath $Folder -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $limit } |
            ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                    Write-Log "Cleanup: gelöscht ($What) -> $($_.FullName)"
                } catch {
                    Write-Log "Cleanup: konnte nicht löschen ($What) -> $($_.FullName): $($_.Exception.Message)" "WARN"
                }
            }
    } catch {
        Write-Log "Cleanup Fehler ($What): $($_.Exception.Message)" "WARN"
    }
}

# Cache (Campaign-Ordner) aufräumen
Cleanup-OldItems -Folder $CacheRoot -KeepDays $cleanupDays -What "Cache"
# Alte Logfiles aufräumen (nur LogonOnce Logs)
Cleanup-OldItems -Folder $LogRoot -KeepDays $cleanupDays -What "Logs"

# =========================
# 4) RUNONCE (HKCU)
# =========================
$RunOnceKey = "HKCU:\Software\Wartungsscript\LogonOnce"
$RunOnceVal = "Campaign_$($cfg.campaignId)"
New-Item -Path $RunOnceKey -Force | Out-Null

$done = (Get-ItemProperty -Path $RunOnceKey -Name $RunOnceVal -ErrorAction SilentlyContinue).$RunOnceVal
if ($done) {
    Write-Log "Campaign bereits erledigt ($RunOnceVal=$done) -> Ende."
    exit 0
}

# =========================
# 5) PRO-CAMPAIGN CACHE (IMMUTABLE)
# =========================
$CampaignCache = Join-Path $CacheRoot $cfg.campaignId
New-Item -Path $CampaignCache -ItemType Directory -Force | Out-Null

function Get-Sha256Hex {
    param([Parameter(Mandatory)][string]$Path)
    try { return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path -ErrorAction Stop).Hash.ToLowerInvariant() }
    catch { return $null }
}

$srcRoot = [string]$cfg.scriptSourceRoot
if (-not $srcRoot) {
    Write-Log "scriptSourceRoot fehlt -> Ende." "ERROR"
    exit 0
}

$ok = 0; $fail = 0
$startAll = Get-Date

foreach ($s in $cfg.subScripts) {

    if (((Get-Date) - $startAll).TotalSeconds -gt $maxRuntime) {
        Write-Log "MaxRuntime ($maxRuntime s) erreicht -> Abbruch." "WARN"
        break
    }

    $name = [string]$s.name
    $file = [string]$s.file
    $args = [string]$s.args
    $expected = ([string]$s.sha256).ToLowerInvariant()

    if (-not $file -or -not $name) {
        Write-Log "Ungültiger subScripts Eintrag (name/file fehlt) -> skip" "WARN"
        $fail++
        if ($stopOnErr) { break }
        continue
    }

    $srcPath = Join-Path $srcRoot $file
    $dstPath = Join-Path $CampaignCache $file

    # Kopieren, wenn noch nicht vorhanden
    if (-not (Test-Path -LiteralPath $dstPath)) {
        try {
            if (-not (Test-Path -LiteralPath $srcPath)) {
                Write-Log "Quelle fehlt: $srcPath -> skip" "WARN"
                $fail++
                if ($stopOnErr) { break }
                continue
            }
            Copy-Item -LiteralPath $srcPath -Destination $dstPath -Force -ErrorAction Stop
            Write-Log "Gecached: $file -> $dstPath"
        } catch {
            Write-Log "Cache Copy Fehler ($file): $($_.Exception.Message)" "ERROR"
            $fail++
            if ($stopOnErr) { break }
            continue
        }
    }

    # Hash prüfen
    if ($expected) {
        $actual = Get-Sha256Hex -Path $dstPath
        if (-not $actual) {
            Write-Log "Hashberechnung fehlgeschlagen ($file) -> skip" "ERROR"
            $fail++
            if ($stopOnErr) { break }
            continue
        }
        if ($actual -ne $expected) {
            Write-Log "HASH MISMATCH ($file). expected=$expected actual=$actual -> skip" "ERROR"
            $fail++
            if ($stopOnErr) { break }
            continue
        }
    } else {
        Write-Log "WARN: Kein sha256 für $file – Ausführung ohne Integritätscheck." "WARN"
    }

    # Ausführen (separater Prozess + Redirect)
    $outFile = Join-Path $CampaignCache ("{0}.stdout.log" -f $name)
    $errFile = Join-Path $CampaignCache ("{0}.stderr.log" -f $name)

    try {
        $argList = @("-NoLogo","-NoProfile","-ExecutionPolicy","Bypass","-File","`"$dstPath`"")
        if ($args -and $args.Trim().Length -gt 0) { $argList += $args }

        Write-Log "Starte: $name ($file) args=[$args]"
        $p = Start-Process powershell.exe -ArgumentList $argList -WindowStyle Hidden `
            -RedirectStandardOutput $outFile -RedirectStandardError $errFile -Wait -PassThru -ErrorAction Stop

        if ($p.ExitCode -eq 0) {
            Write-Log "OK: $name ExitCode=0"
            $ok++
        } else {
            Write-Log "FAIL: $name ExitCode=$($p.ExitCode)" "WARN"
            $fail++
            if ($stopOnErr) { break }
        }
    } catch {
        Write-Log "ERROR: $name -> $($_.Exception.Message)" "ERROR"
        $fail++
        if ($stopOnErr) { break }
    }
}

Write-Log "Subskripte fertig. OK=$ok FAIL=$fail"

# =========================
# 6) RUNONCE FLAG SETZEN
# =========================
try {
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    New-ItemProperty -Path $RunOnceKey -Name $RunOnceVal -Value $stamp -PropertyType String -Force | Out-Null
    Write-Log "Done gesetzt: $RunOnceVal=$stamp"
} catch {
    Write-Log "Done Flag konnte nicht gesetzt werden: $($_.Exception.Message)" "WARN"
}

# =========================
# 7) OPTIONAL: ZENTRALER LOG UPLOAD
# =========================
try {
    $central = [string]$cfg.centralLogRoot
    if ($central -and $central.Trim().Length -gt 0 -and (Test-Path $central)) {
        $target = Join-Path $central $env:COMPUTERNAME
        $target = Join-Path $target $env:USERNAME
        $target = Join-Path $target ("LogonOnce_{0}" -f $cfg.campaignId)
        New-Item -Path $target -ItemType Directory -Force | Out-Null

        Copy-Item -LiteralPath $MasterLog -Destination (Join-Path $target (Split-Path $MasterLog -Leaf)) -Force -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath $CampaignCache -Destination (Join-Path $target "CacheLogs") -Recurse -Force -ErrorAction SilentlyContinue

        Write-Log "Zentraler Logcopy versucht: $target"
    }
} catch {
    Write-Log "Zentraler Logcopy Fehler: $($_.Exception.Message)" "WARN"
}

Write-Log "LogonOnce beendet."
exit 0