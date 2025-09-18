# Speedtest-Logger.ps1 — SecureNet NetMon (PS 5.1) with PingPlotter-style ICMP + improved DNS timing
$ErrorActionPreference = 'Stop'

# ===== Settings =====
$BaseDir   = 'C:\ProgramData\SecureNet\NetMon'
$LogDir    = Join-Path $BaseDir 'logs'
$JsonDir   = Join-Path $LogDir  'json'
$PcapDir   = Join-Path $LogDir  'pcap'
$PingDir   = Join-Path $BaseDir 'pings'
$BinPath   = 'C:\Program Files\Speedtest\speedtest.exe'
$ServerId  = 59734
$IfaceName = $env:INTERFACE_NAME

$RetryThresholdMbps = 50
$RetryDelaySeconds  = 5
$LossFlagPct        = 1.0
$PktmonSeconds      = 10
$KeepDays           = 14

$EnableEventLog = $true
$EventSource    = 'SecureNet-NetMon'
$EventLogName   = 'Application'

New-Item -ItemType Directory -Force -Path $LogDir,$JsonDir,$PcapDir | Out-Null

# ---- Site tags (optional) ----
$SiteCfgPath = Join-Path $BaseDir 'site.json'
$Customer=''; $Site=''; $Device=''
if (Test-Path $SiteCfgPath) {
  try { $siteCfg = Get-Content $SiteCfgPath -Raw | ConvertFrom-Json } catch {}
  if ($siteCfg) {
    if ($siteCfg.customer) { $Customer = [string]$siteCfg.customer }
    if ($siteCfg.site)     { $Site     = [string]$siteCfg.site }
    if ($siteCfg.device)   { $Device   = [string]$siteCfg.device }
  }
}

# ===== Helpers =====
function Save-RawJson($text, $dir) {
  $stamp = Get-Date; $ymd=$stamp.ToString('yyyy-MM-dd'); $hms=$stamp.ToString('HHmmss')
  $outDir = Join-Path $dir $ymd; New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  $path = Join-Path $outDir "$hms.json"; if ($text) { $text | Out-File -Encoding UTF8 $path }; return $path
}
function Run-Speedtest([bool]$withoutServerId = $false) {
  $args = @('--format=json','--progress=no','--accept-license','--accept-gdpr')
  if (-not $withoutServerId -and $ServerId) { $args += @('--server-id', $ServerId) }
  try { & "$BinPath" @args 2>&1 } catch { "" }
}
function Parse-Speedtest($jsonText) {
  $o=$null; try { $o=$jsonText | ConvertFrom-Json -ErrorAction Stop } catch { return @{ parsed=$null; error=($jsonText -replace '\s+',' ') } }
  return @{ parsed=$o; error=$null }
}
function Mbps-From($bandwidthBytesPerSec) { if ($bandwidthBytesPerSec -eq $null) { return $null }; [math]::Round(($bandwidthBytesPerSec*8.0)/1000000.0,2) }
function Derive-Mbps($bytes,$elapsedMs) { try { if ($bytes -and $elapsedMs -gt 0) { return [math]::Round(((8.0*[double]$bytes)/([double]$elapsedMs/1000.0))/1000000.0,2) } } catch {}; $null }
function Test-AvgPing($target,$count=4,$timeout=2000) {
  try { $res = Test-Connection -ComputerName $target -Count $count -TimeoutSeconds ([math]::Ceiling($timeout/1000)) -ErrorAction Stop
        if ($res) { return [math]::Round(($res | Measure-Object -Property ResponseTime -Average).Average, 2) } } catch {}
  return $null
}
# Better DNS timing: use nslookup against first configured DNS, random label to avoid cache, 2s timeout
function Get-FirstDnsServer() {
  try {
    $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' | Sort-Object RouteMetric,InterfaceMetric | Select-Object -First 1
    if ($route) {
      $idx = $route.ifIndex
      $dns = (Get-DnsClientServerAddress -InterfaceIndex $idx -AddressFamily IPv4 -ErrorAction Stop).ServerAddresses
      if ($dns -and $dns.Count -gt 0) { return $dns[0] }
    }
  } catch {}
  return '1.1.1.1'
}

# Robust DNS timing: try Resolve-DnsName with timeout; fall back to nslookup. Random label avoids cache.
function Measure-DnsLookup {
  param([string]$qName = 'microsoft.com', [int]$timeoutSec = 2)

  function Get-FirstDnsServer {
    try {
      $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' | Sort-Object RouteMetric,InterfaceMetric | Select-Object -First 1
      if ($route) {
        $idx = $route.ifIndex
        $dns = (Get-DnsClientServerAddress -InterfaceIndex $idx -AddressFamily IPv4 -ErrorAction Stop).ServerAddresses
        if ($dns -and $dns.Count -gt 0) { return $dns[0] }
      }
    } catch {}
    return '1.1.1.1'
  }

  $resolver = Get-FirstDnsServer
  $rand = -join ((48..57 + 97..122) | Get-Random -Count 8 | ForEach-Object {[char]$_})
  $name = "$rand.$qName"

  # Try Resolve-DnsName in a background job to enforce timeout
  try {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $job = Start-Job -ScriptBlock {
      param($n,$server)
      try { Resolve-DnsName -Name $n -Server $server -Type A -DnsOnly -NoHostsFile -ErrorAction Stop | Out-Null; return 0 }
      catch { return 1 }
    } -ArgumentList $name,$resolver

    if (Wait-Job -Job $job -Timeout $timeoutSec | Out-Null) {
      $code = Receive-Job $job -ErrorAction SilentlyContinue
      Remove-Job $job -Force -ErrorAction SilentlyContinue
      $sw.Stop()
      return [math]::Round($sw.Elapsed.TotalMilliseconds,2)
    } else {
      Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
  } catch {}

  # Fallback to nslookup (capture stdout+stderr, hidden window)
  try {
    $tmp = [System.IO.Path]::GetTempFileName()
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $p = Start-Process -FilePath "$env:SystemRoot\System32\nslookup.exe" `
         -ArgumentList @("-timeout=$timeoutSec","-type=A",$name,$resolver) `
         -WindowStyle Hidden -PassThru -Wait `
         -RedirectStandardOutput $tmp -RedirectStandardError $tmp
    $sw2.Stop()
    try { Remove-Item $tmp -ErrorAction SilentlyContinue } catch {}
    return [math]::Round($sw2.Elapsed.TotalMilliseconds,2)
  } catch { return $null }
}

function Start-PktmonCapture($seconds, [string]$outDir) {
  $ts=Get-Date; $base=("cap-{0:yyyyMMdd-HHmmss}" -f $ts); $etl=Join-Path $outDir "$base.etl"; $pcap=Join-Path $outDir "$base.pcapng"
  try { Start-Process -FilePath pktmon.exe -ArgumentList @('start','--capture','--file',"$etl") -WindowStyle Hidden -NoNewWindow
        Start-Sleep -Seconds ([int]$seconds); pktmon.exe stop | Out-Null; pktmon.exe format "$etl" -o "$pcap" | Out-Null; return $pcap } catch { return $null }
}
# Aggregate last window of PingSampler data
function Summarize-Pings([datetime]$from,[datetime]$to) {
  $out = @{ target=''; samples=0; success=0; lossPct=$null; avg=$null; min=$null; max=$null; jitter=$null }
  if (-not (Test-Path $PingDir)) { return $out }
  $dates = @($from.Date, $to.Date) | Sort-Object -Unique
  $rows = @()
  foreach($d in $dates) {
    $f = Join-Path $PingDir ($d.ToString('yyyy-MM-dd') + '.csv')
    if (Test-Path $f) {
      $rows += (Import-Csv $f | Where-Object {
        $_.timestamp -and (Get-Date $_.timestamp) -ge $from -and (Get-Date $_.timestamp) -le $to
      })
    }
  }
  if (-not $rows -or $rows.Count -eq 0) { return $out }
  $out.target  = $rows[0].target
  $out.samples = $rows.Count
  $ok = 0; $lat = @()
  foreach($r in $rows) {
    if ($r.success -eq '1') {
      $ok++
      if ($r.rttMs -ne $null -and "$($r.rttMs)" -ne '') { $lat += ([double]$r.rttMs) }
    }
  }
  $out.success = $ok
  $out.lossPct = [math]::Round(((($rows.Count - $ok) * 100.0) / $rows.Count),2)
  if ($lat.Count -gt 0) {
    $out.min = [math]::Round(($lat | Measure-Object -Minimum).Minimum,2)
    $out.max = [math]::Round(($lat | Measure-Object -Maximum).Maximum,2)
    $out.avg = [math]::Round(($lat | Measure-Object -Average).Average,2)
    # mean absolute successive difference as simple jitter proxy
    if ($lat.Count -gt 1) {
      $diffs = @()
      for($i=1;$i -lt $lat.Count;$i++){ $diffs += [math]::Abs($lat[$i]-$lat[$i-1]) }
      $out.jitter = [math]::Round((($diffs | Measure-Object -Average).Average),2)
    }
  }
  return $out
}

# Ensure Event Log
if ($EnableEventLog) { try { if (-not [System.Diagnostics.EventLog]::SourceExists($EventSource)) { New-EventLog -LogName $EventLogName -Source $EventSource } } catch {} }

# ===== Local path health (gateway/DNS) =====
$gwRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Sort-Object RouteMetric,InterfaceMetric | Select-Object -First 1
$gateway = $null; if ($gwRoute) { $gateway = $gwRoute.NextHop }
$gatewayPing = if ($gateway) { Test-AvgPing $gateway } else { $null }
$dnsLookupMs = Measure-DnsLookup 'google.com'  # uses first DNS server, random label to avoid cache

# ===== Speedtest runs =====
$stamp    = Get-Date
$json1    = Run-Speedtest
Save-RawJson $json1 $JsonDir | Out-Null
$p        = Parse-Speedtest $json1
$pObj     = $p.parsed
$errorMsg = $p.error

$downloadMbps = $null; $uploadMbps = $null
$latencyMs    = $null; $jitterMs   = $null
$packetLoss   = $null
$serverName   = $null; $serverLoc  = $null; $serverIdOut = $null
$extIp        = $null; $isp        = $null
$resultId     = $null; $resultUrl  = $null

if ($pObj) {
  $downloadMbps = Mbps-From $pObj.download.bandwidth; if ($downloadMbps -eq $null) { $downloadMbps = Derive-Mbps $pObj.download.bytes $pObj.download.elapsed }
  $uploadMbps   = Mbps-From $pObj.upload.bandwidth;   if ($uploadMbps   -eq $null) { $uploadMbps   = Derive-Mbps $pObj.upload.bytes   $pObj.upload.elapsed   }
  $latencyMs    = $pObj.ping.latency; $jitterMs = $pObj.ping.jitter
  $packetLoss   = $pObj.packetLoss
  $serverName   = $pObj.server.name; $serverLoc = $pObj.server.location; $serverIdOut = $pObj.server.id
  $extIp        = $pObj.interface.externalIp; $isp = $pObj.isp
  $resultId     = $pObj.result.id; $resultUrl = $pObj.result.url
  if ($pObj.error) { $errorMsg = "$($pObj.error)" }
}

# Retry if degraded
$didRetry=$false; $pcapPath=$null
$degraded = (($downloadMbps -eq $null) -or ($downloadMbps -lt $RetryThresholdMbps) -or ($packetLoss -ne $null -and $packetLoss -gt $LossFlagPct) -or $errorMsg)
if ($degraded) {
  Start-Sleep -Seconds $RetryDelaySeconds
  $didRetry=$true
  $json2 = Run-Speedtest
  Save-RawJson $json2 $JsonDir | Out-Null
  $p2 = Parse-Speedtest $json2; $q=$p2.parsed
  if ($q) {
    $dl2 = Mbps-From $q.download.bandwidth; if ($dl2 -eq $null) { $dl2 = Derive-Mbps $q.download.bytes $q.download.elapsed }
    $ul2 = Mbps-From $q.upload.bandwidth;   if ($ul2 -eq $null) { $ul2 = Derive-Mbps $q.upload.bytes   $q.upload.elapsed   }
    if ($dl2 -ne $null) { $downloadMbps=$dl2 }; if ($ul2 -ne $null) { $uploadMbps=$ul2 }
    if ($q.ping.latency -ne $null) { $latencyMs=$q.ping.latency }
    if ($q.ping.jitter  -ne $null) { $jitterMs =$q.ping.jitter  }
    if ($q.packetLoss   -ne $null) { $packetLoss=$q.packetLoss  }
    if ($q.server.name) { $serverName=$q.server.name }; if ($q.server.location) { $serverLoc=$q.server.location }
    if ($q.server.id) { $serverIdOut=$q.server.id }
    if ($q.interface.externalIp) { $extIp = $q.interface.externalIp }
    if ($q.isp) { $isp = $q.isp }
    if ($q.result.id)  { $resultId=$q.result.id }
    if ($q.result.url) { $resultUrl=$q.result.url }
    if ($q.error) { $errorMsg="$($q.error)" } else { $errorMsg=$null }
  }
  if ((($downloadMbps -eq $null) -or ($downloadMbps -lt $RetryThresholdMbps) -or ($packetLoss -ne $null -and $packetLoss -gt $LossFlagPct) -or $errorMsg)) {
    $pcapPath = Start-PktmonCapture -seconds $PktmonSeconds -outDir $PcapDir
    if ($EnableEventLog) {
      try {
        $msg = ("Speedtest degraded. DL={0} Mbps, UL={1} Mbps, Lat={2} ms, Loss={3}% | Server {4} {5} | pcap={6}" -f `
                $downloadMbps,$uploadMbps,$latencyMs,$packetLoss,$serverIdOut,$serverName,$pcapPath)
        Write-EventLog -LogName $EventLogName -Source $EventSource -EntryType Warning -EventId 4001 -Message $msg
      } catch {}
    }
  }
}

# ===== Summarize last 15 minutes of ICMP samples =====
$windowEnd   = Get-Date
$windowStart = $windowEnd.AddMinutes(-15)
$icmp = Summarize-Pings -from $windowStart -to $windowEnd  # returns empty/default if sampler not running

# ===== Write CSV =====
$ymd = (Get-Date).ToString('yyyy-MM-dd')
$csvPath = Join-Path $LogDir "$ymd-speedtests.csv"
$exists  = Test-Path $csvPath
if (-not $exists) {
  # original cols + previous additions + new ICMP summary (at end)
  "timestamp,extIp,isp,ifaceName,downloadMbps,uploadMbps,latencyMs,jitterMs,packetLossPct,serverName,serverLocation,serverId,resultId,error,resultUrl,didRetry,gatewayPingMs,dnsLookupMs,pcapPath,customer,site,device,icmpTarget,icmpSamples,icmpSuccess,icmpLossPct,icmpAvgMs,icmpMinMs,icmpMaxMs,icmpJitterMs" |
    Out-File -Encoding UTF8 $csvPath
}

$line = [pscustomobject]@{
  timestamp      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  extIp          = $extIp
  isp            = $isp
  ifaceName      = $IfaceName
  downloadMbps   = $downloadMbps
  uploadMbps     = $uploadMbps
  latencyMs      = $latencyMs
  jitterMs       = $jitterMs
  packetLossPct  = $packetLoss
  serverName     = $serverName
  serverLocation = $serverLoc
  serverId       = $serverIdOut
  resultId       = $resultId
  error          = $errorMsg
  resultUrl      = $resultUrl
  didRetry       = $didRetry
  gatewayPingMs  = $gatewayPing
  dnsLookupMs    = $dnsLookupMs
  pcapPath       = $pcapPath
  customer       = $Customer
  site           = $Site
  device         = $Device
  icmpTarget     = $icmp.target
  icmpSamples    = $icmp.samples
  icmpSuccess    = $icmp.success
  icmpLossPct    = $icmp.lossPct
  icmpAvgMs      = $icmp.avg
  icmpMinMs      = $icmp.min
  icmpMaxMs      = $icmp.max
  icmpJitterMs   = $icmp.jitter
}

$line | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -Append -Encoding UTF8 $csvPath

# ===== Retention =====
try {
  Get-ChildItem $LogDir -Filter '*-speedtests.csv' -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } | Remove-Item -Force -ErrorAction SilentlyContinue
  if (Test-Path $JsonDir) { Get-ChildItem $JsonDir -Recurse -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } | Remove-Item -Force -ErrorAction SilentlyContinue }
  if (Test-Path $PcapDir) { Get-ChildItem $PcapDir -Recurse -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } | Remove-Item -Force -ErrorAction SilentlyContinue }
} catch {}
