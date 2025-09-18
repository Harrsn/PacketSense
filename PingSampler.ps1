# PingSampler.ps1 — 1Hz ICMP sampler with self-logging (PS 5.1)
$ErrorActionPreference = 'SilentlyContinue'

$BaseDir = 'C:\ProgramData\SecureNet\NetMon'
$PingDir = Join-Path $BaseDir 'pings'
$LogFile = Join-Path $PingDir 'PingSampler.log'
$CfgPath = Join-Path $BaseDir 'config.json'
$Target  = '8.8.8.8'  # default; override via config.json { "pingTarget": "23.186.16.138" }

New-Item -ItemType Directory -Force -Path $PingDir | Out-Null
if (Test-Path $CfgPath) {
  try { $cfg = Get-Content $CfgPath -Raw | ConvertFrom-Json; if ($cfg -and $cfg.pingTarget) { $Target = [string]$cfg.pingTarget } } catch {}
}

function Log([string]$msg){ $ts=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; Add-Content -Path $LogFile -Value "$ts $msg" }
Log "PingSampler starting; target=$Target; PS=$($PSVersionTable.PSVersion)"

$lastHeartbeat = Get-Date

while ($true) {
  $stamp = Get-Date
  $ymd   = $stamp.ToString('yyyy-MM-dd')
  $csv   = Join-Path $PingDir "$ymd.csv"
  if (-not (Test-Path $csv)) {
    "timestamp,target,success,rttMs" | Out-File -Encoding UTF8 $csv
    Log "New CSV created: $csv"
  }

  $ok=0; $rtt=$null
  try {
    $res = Test-Connection -ComputerName $Target -Count 1 -IPv4 -ErrorAction Stop
    if ($res) { $ok=1; $rtt=[math]::Round(($res | Select-Object -ExpandProperty ResponseTime),2) }
  } catch { $ok=0 }

  "{0},{1},{2},{3}" -f $stamp.ToString('yyyy-MM-dd HH:mm:ss'),$Target,$ok,($rtt -as [string]) | Add-Content $csv

  if ((Get-Date) -ge $lastHeartbeat.AddMinutes(1)) {
    Log "alive: last=$($stamp.ToString('HH:mm:ss')) ok=$ok rtt=$rtt"
    $lastHeartbeat = Get-Date
  }

  Start-Sleep -Milliseconds 1000
}
