param(
  [switch]$Once
)

$ErrorActionPreference = 'SilentlyContinue'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptRoot 'config.json'
$StateDir = Join-Path $ScriptRoot 'state'
$StateFile = Join-Path $StateDir 'version.json'

function Ensure-Dir($p) { if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null } }

function Load-Config {
  if (-not (Test-Path $ConfigPath)) { throw "config.json not found" }
  return (Get-Content $ConfigPath -Raw | ConvertFrom-Json)
}

function Write-Log($cfg, [string]$msg) {
  $logPath = Join-Path $ScriptRoot $cfg.logging.path
  Ensure-Dir (Split-Path -Parent $logPath)
  $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
  Add-Content -Path $logPath -Value $line -Encoding UTF8

  $maxBytes = [int64]$cfg.logging.maxSizeMB * 1MB
  if ((Test-Path $logPath) -and ((Get-Item $logPath).Length -gt $maxBytes)) {
    $bak = "$logPath.bak"
    if (Test-Path $bak) { Remove-Item $bak -Force }
    Move-Item $logPath $bak -Force
  }
}

function Get-Metrics {
  $os = Get-CimInstance Win32_OperatingSystem
  $cpu = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples[0].CookedValue
  $memUsedPct = [math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 1)
  $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
  $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)

  [pscustomobject]@{
    CpuPercent = [math]::Round($cpu,1)
    MemoryPercent = $memUsedPct
    SystemDriveFreeGB = $freeGB
  }
}

function Remove-OldFiles([string]$path, [int]$olderThanDays, [int64]$maxDeleteBytes) {
  if (-not (Test-Path $path)) { return 0 }
  $cutoff = (Get-Date).AddDays(-$olderThanDays)
  $deleted = 0

  Get-ChildItem -Path $path -Recurse -File -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt $cutoff } |
    Sort-Object Length -Descending |
    ForEach-Object {
      if ($deleted -ge $maxDeleteBytes) { return }
      $len = $_.Length
      Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
      if (-not (Test-Path $_.FullName)) { $deleted += $len }
    }

  return $deleted
}

function Invoke-Cleanup($cfg) {
  $totalDeleted = 0
  $maxDeleteBytes = [int64]$cfg.actions.maxDeleteMBPerRun * 1MB

  if ($cfg.actions.cleanupTemp) {
    $temp1 = $env:TEMP
    $temp2 = "$env:WINDIR\Temp"
    $totalDeleted += Remove-OldFiles -path $temp1 -olderThanDays $cfg.actions.tempFileOlderThanDays -maxDeleteBytes $maxDeleteBytes
    $remaining = [math]::Max(0, $maxDeleteBytes - $totalDeleted)
    if ($remaining -gt 0) {
      $totalDeleted += Remove-OldFiles -path $temp2 -olderThanDays $cfg.actions.tempFileOlderThanDays -maxDeleteBytes $remaining
    }
  }

  if ($cfg.actions.cleanupDeliveryOptimization) {
    $doPath = "$env:ProgramData\Microsoft\Windows\DeliveryOptimization\Cache"
    $remaining = [math]::Max(0, $maxDeleteBytes - $totalDeleted)
    if ($remaining -gt 0) {
      $totalDeleted += Remove-OldFiles -path $doPath -olderThanDays 2 -maxDeleteBytes $remaining
    }
  }

  return [math]::Round($totalDeleted / 1MB, 2)
}

function Invoke-TrimWorkingSet($cfg) {
  if (-not $cfg.actions.trimWorkingSet) { return 0 }

  Add-Type @"
using System;
using System.Runtime.InteropServices;
public class MemUtil {
  [DllImport("psapi.dll")] public static extern int EmptyWorkingSet(IntPtr hwProc);
}
"@

  $count = 0
  foreach ($name in $cfg.actions.trimProcessAllowlist) {
    Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
      [void][MemUtil]::EmptyWorkingSet($_.Handle)
      $count++
    }
  }
  return $count
}

function Load-State {
  Ensure-Dir $StateDir
  if (Test-Path $StateFile) {
    return (Get-Content $StateFile -Raw | ConvertFrom-Json)
  }
  return [pscustomobject]@{ lastUpdateCheck = "1970-01-01T00:00:00Z"; version = "1.0.0" }
}

function Save-State($s) {
  $s | ConvertTo-Json | Set-Content -Path $StateFile -Encoding UTF8
}

function Invoke-SelfUpdate($cfg) {
  if (-not $cfg.selfUpdate.enabled) { return "skip" }

  $state = Load-State
  $last = Get-Date $state.lastUpdateCheck
  $due = ((Get-Date) - $last).TotalHours -ge [double]$cfg.selfUpdate.checkEveryHours
  if (-not $due) { return "not-due" }

  $state.lastUpdateCheck = (Get-Date).ToString("o")
  Save-State $state

  try {
    $manifest = Invoke-RestMethod -Uri $cfg.selfUpdate.manifestUrl -Method Get -TimeoutSec 10
    if (-not $manifest.version -or -not $manifest.downloadUrl -or -not $manifest.sha256) { return "bad-manifest" }

    if ($manifest.version -eq $state.version) { return "up-to-date" }

    $tmp = Join-Path $env:TEMP "host-perf-optimizer-update.ps1"
    Invoke-WebRequest -Uri $manifest.downloadUrl -OutFile $tmp -UseBasicParsing -TimeoutSec 20
    $hash = (Get-FileHash -Path $tmp -Algorithm SHA256).Hash.ToLower()
    if ($hash -ne ($manifest.sha256.ToLower())) { Remove-Item $tmp -Force; return "hash-mismatch" }

    Copy-Item $tmp (Join-Path $ScriptRoot 'optimizer.ps1') -Force
    Remove-Item $tmp -Force

    $state.version = $manifest.version
    Save-State $state
    return "updated:$($manifest.version)"
  } catch {
    return "update-error"
  }
}

$cfg = Load-Config
$metrics = Get-Metrics
Write-Log $cfg "metrics cpu=$($metrics.CpuPercent)% mem=$($metrics.MemoryPercent)% diskFree=$($metrics.SystemDriveFreeGB)GB"

$needAct = ($metrics.CpuPercent -ge $cfg.thresholds.cpuPercent) -or
           ($metrics.MemoryPercent -ge $cfg.thresholds.memoryPercent) -or
           ($metrics.SystemDriveFreeGB -le $cfg.thresholds.systemDriveFreeGB)

if ($needAct) {
  $freedMB = Invoke-Cleanup $cfg
  $trimmed = Invoke-TrimWorkingSet $cfg
  Write-Log $cfg "optimize triggered freed=${freedMB}MB trimmedProc=$trimmed"
}

$up = Invoke-SelfUpdate $cfg
Write-Log $cfg "self-update=$up"

if (-not $Once) {
  Start-Sleep -Seconds 1
}
