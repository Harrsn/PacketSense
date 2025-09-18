# Build-Speedtest-Report.ps1 — SVG-only report (no JS), with ICMP (PingSampler) support
$ErrorActionPreference = 'Stop'

# ===== Paths / window =====
$BaseDir = 'C:\ProgramData\SecureNet\NetMon'
$LogDir  = Join-Path $BaseDir 'logs'
$RptDir  = Join-Path $BaseDir 'reports'
$OutHtml = Join-Path $RptDir 'NetReport.html'
$DaysBack = 7
New-Item -ItemType Directory -Force -Path $RptDir | Out-Null

# ===== SLA / plan thresholds (tune these) =====
$PlanDlMbps        = 300    # expected min download
$PlanUlMbps        = 20     # expected min upload
$LatencyMaxMs      = 40     # max acceptable latency
$LossMaxPct        = 1.0    # max acceptable packet loss (preferred ICMP)
$GatewayPingMaxMs  = 10     # gateway ping ok if <= this
$DnsLookupMaxMs    = 100    # DNS lookup ok if <= this

# ===== Helpers =====
function Fmt($x){ if ($null -eq $x -or "$x" -eq '') { 'n/a' } else { "$x" } }
function To-Double($x){
  if ($null -eq $x -or "$x" -eq '') { return $null }
  $s = "$x" -replace ',', '.'
  $s = ($s -replace '[^0-9\.\-]', '')
  if ($s -in @('','.','-')) { return $null }
  try { return [double]$s } catch { return $null }
}
function Parse-Ts($t){
  if ($null -eq $t -or "$t" -eq '') { return $null }
  $s = "$t" -replace 'T',' '
  try { return [datetime]$s } catch { return $null }
}
function Stats($arr){
  $vals = @()
  foreach($v in $arr){ if ($null -ne $v) { try { $d=[double]$v; if(-not [double]::IsNaN($d) -and -not [double]::IsInfinity($d)){ $vals += $d } } catch {} } }
  if ($vals.Count -eq 0) { return @{min=$null; avg=$null; max=$null} }
  return @{
    min = [math]::Round(($vals | Measure-Object -Minimum).Minimum,2)
    avg = [math]::Round(($vals | Measure-Object -Average).Average,2)
    max = [math]::Round(($vals | Measure-Object -Maximum).Maximum,2)
  }
}
function HtmlEscape([object]$s){
  if ($null -eq $s) { return '' }
  $t = "$s"
  $t = $t -replace '&','&amp;'
  $t = $t -replace '<','&lt;'
  $t = $t -replace '>','&gt;'
  return $t
}
function GetProp($obj,[string]$name){ if ($null -eq $obj) { return $null }; if ($obj.PSObject.Properties.Name -contains $name) { return $obj.$name } else { return $null } }

# SVG chart (multi-line + optional threshold)
function New-SvgChart {
 param(
   [string]$Title,
   [hashtable[]]$Series,      # @{ name='Download'; values=@(...); stroke='#4FC3F7' }
   [string]$YUnit='',
   [double]$Threshold=$null,  # draw a horizontal threshold line if provided
   [int]$Width=760,
   [int]$Height=260,
   [int]$Pad=40
 )
  $plotW = $Width  - ($Pad*2)
  $plotH = $Height - ($Pad*2)

  # Determine X count
  $count = 0
  foreach ($s in $Series) { $n = ($s.values).Count; if ($n -gt $count) { $count = $n } }
  if ($count -lt 2) {
    return "<svg width='$Width' height='$Height' xmlns='http://www.w3.org/2000/svg'>
      <rect x='0' y='0' width='$Width' height='$Height' fill='#121923' stroke='#1c2532'/>
      <text x='50%' y='50%' fill='#e6eef6' font-family='Segoe UI, Arial' font-size='14' text-anchor='middle'>No data to plot</text>
    </svg>"
  }

  # Y max from all numeric values + threshold
  $allVals = @()
  foreach ($s in $Series) { $allVals += ($s.values | Where-Object { $_ -ne $null }) }
  if ($Threshold -ne $null) { $allVals += $Threshold }
  if (-not $allVals) { $allVals = @(0) }
  $maxVal = [double](($allVals | Measure-Object -Maximum).Maximum)
  if (-not $maxVal -or $maxVal -le 0) { $maxVal = 1 }
  $scaled = [math]::Ceiling($maxVal * 1.1)
  $yMax   = [double]([math]::Max($scaled, 1))

  # Grid (5 lines)
  $grid = ""
  for ($i=0; $i -le 5; $i++) {
    $y = $Pad + $plotH - [math]::Round($plotH * ($i/5.0))
    $val = [math]::Round($yMax * ($i/5.0),2)
    $grid += "<line x1='$Pad' y1='$y' x2='$( $Pad+$plotW )' y2='$y' stroke='rgba(230,238,246,0.12)' stroke-width='1'/>"
    $grid += "<text x='5' y='$( $y+4 )' fill='#e6eef6' font-size='11' font-family='Segoe UI, Arial'>$val</text>"
  }

  # Axes & title
  $axes = "<rect x='$Pad' y='$Pad' width='$plotW' height='$plotH' fill='none' stroke='rgba(230,238,246,0.15)'/>"
  $titleText = if ($YUnit) { "$Title ($YUnit)" } else { $Title }
  $yTitle = "<text x='$Pad' y='$( $Pad-10 )' fill='#e6eef6' font-size='12' font-family='Segoe UI, Arial'>$titleText</text>"

  # Threshold line
  $th = ""
  if ($Threshold -ne $null) {
    $thY = $Pad + $plotH - [math]::Round(($plotH) * ([double]$Threshold / $yMax))
    $th  = "<line x1='$Pad' y1='$thY' x2='$( $Pad+$plotW )' y2='$thY' stroke='rgba(255,255,255,0.35)' stroke-dasharray='4 4' stroke-width='1'/>" +
           "<text x='$( $Pad+$plotW-4 )' y='$( $thY-4 )' text-anchor='end' fill='#e6eef6' font-size='11' font-family='Segoe UI, Arial'>threshold: $Threshold</text>"
  }

  # Lines + points
  $lines = ""
  foreach ($s in $Series) {
    $vals = $s.values
    $pts = @()
    for ($i=0; $i -lt $count; $i++) {
      $v = if ($i -lt $vals.Count) { $vals[$i] } else { $null }
      if ($v -eq $null) { continue }
      $x = $Pad + [math]::Round(($plotW) * ($i / [math]::Max($count-1,1)))
      $y = $Pad + $plotH - [math]::Round(($plotH) * ([double]$v / $yMax))
      $pts += "$x,$y"
    }
    if ($pts.Count -ge 2) {
      $stroke = if ($s.ContainsKey('stroke')) { $s.stroke } else { '#4FC3F7' }
      $lines += "<polyline fill='none' stroke='$stroke' stroke-width='2' points='" + ($pts -join ' ') + "'/>"
      foreach ($pt in $pts) { $xy = $pt -split ','; $lines += "<circle cx='$($xy[0])' cy='$($xy[1])' r='2' fill='$stroke'/>" }
    }
  }

  @"
<svg width='$Width' height='$Height' xmlns='http://www.w3.org/2000/svg'>
  <rect x='0' y='0' width='$Width' height='$Height' fill='#121923' stroke='#1c2532'/>
  $yTitle
  $grid
  $th
  $axes
  $lines
</svg>
"@
}

# Heat strip (overall health per run)
function New-HeatStrip {
 param(
   [int]$Count,
   [string[]]$Colors, # array of color hex per run index
   [int]$Width=760,
   [int]$Height=20,
   [int]$Pad=4
 )
  if ($Count -lt 1) {
    return "<svg width='$Width' height='$Height' xmlns='http://www.w3.org/2000/svg'>
      <rect x='0' y='0' width='$Width' height='$Height' fill='#121923' stroke='#1c2532'/>
      <text x='50%' y='50%' fill='#e6eef6' font-family='Segoe UI, Arial' font-size='12' text-anchor='middle' dominant-baseline='middle'>No data</text>
    </svg>"
  }
  $plotW = $Width - ($Pad*2)
  $w = [math]::Max([math]::Floor($plotW / $Count), 2)
  $x0 = $Pad
  $rects = ""
  for ($i=0; $i -lt $Count; $i++) {
    $x = $x0 + ($i * $w)
    $c = if ($i -lt $Colors.Count) { $Colors[$i] } else { '#666' }
    $rects += "<rect x='$x' y='$Pad' width='$w' height='$( $Height - $Pad*2 )' fill='$c'/>"
  }
  @"
<svg width='$Width' height='$Height' xmlns='http://www.w3.org/2000/svg'>
  <rect x='0' y='0' width='$Width' height='$Height' fill='#121923' stroke='#1c2532'/>
  $rects
</svg>
"@
}

# ===== Load & normalize data =====
$cutoff = (Get-Date).AddDays(-$DaysBack)
$files  = Get-ChildItem -Path $LogDir -Filter '*-speedtests.csv' -ErrorAction SilentlyContinue |
          Where-Object { $_.LastWriteTime -ge $cutoff }

if (-not $files) {
  "<html><body style='font-family:Segoe UI,Arial'><h2>No speedtest logs found in $LogDir for last $DaysBack day(s).</h2></body></html>" |
    Out-File -Encoding UTF8 $OutHtml; return
}

$data = $files | ForEach-Object { Import-Csv $_.FullName -ErrorAction SilentlyContinue } |
  Where-Object { $_.timestamp } |
  ForEach-Object {
    $ts = Parse-Ts $_.timestamp; if ($null -eq $ts) { return }
    [pscustomobject]@{
      ts        = $ts
      dlMbps    = To-Double (GetProp $_ 'downloadMbps')
      ulMbps    = To-Double (GetProp $_ 'uploadMbps')
      latMs     = To-Double (GetProp $_ 'latencyMs')
      jitMs     = To-Double (GetProp $_ 'jitterMs')
      # prefer ICMP (PingSampler); keep CLI packetLoss for fallback
      icmpLoss  = To-Double (GetProp $_ 'icmpLossPct')
      lossPct   = To-Double (GetProp $_ 'packetLossPct')
      # ICMP ping stats
      icmpAvg   = To-Double (GetProp $_ 'icmpAvgMs')
      icmpMin   = To-Double (GetProp $_ 'icmpMinMs')
      icmpMax   = To-Double (GetProp $_ 'icmpMaxMs')
      icmpJit   = To-Double (GetProp $_ 'icmpJitterMs')
      # path health
      gwMs      = To-Double (GetProp $_ 'gatewayPingMs')
      dnsMs     = To-Double (GetProp $_ 'dnsLookupMs')
      server    = GetProp $_ 'serverName'
      serverId  = GetProp $_ 'serverId'
      loc       = GetProp $_ 'serverLocation'
      isp       = GetProp $_ 'isp'
      extIp     = GetProp $_ 'extIp'
      retry     = GetProp $_ 'didRetry'
      url       = GetProp $_ 'resultUrl'
      error     = GetProp $_ 'error'
    }
  } | Where-Object { $_.ts -ge $cutoff } | Sort-Object ts

if (-not $data) {
  "<html><body style='font-family:Segoe UI,Arial'><h2>No parsable rows in CSV for last $DaysBack day(s).</h2></body></html>" |
    Out-File -Encoding UTF8 $OutHtml; return
}

# ===== Summaries =====
$rows     = $data.Count
$firstTs  = $data[0].ts
$lastTs   = $data[-1].ts
$dlS      = Stats ($data | Select-Object -ExpandProperty dlMbps)
$ulS      = Stats ($data | Select-Object -ExpandProperty ulMbps)
$ltS      = Stats ($data | Select-Object -ExpandProperty latMs)
$jtS      = Stats ($data | Select-Object -ExpandProperty jitMs)

# Preferred loss = ICMP if present, else CLI packetLoss
$prefLossVals = @()
foreach($r in $data){ $prefLossVals += ($(if ($r.icmpLoss -ne $null) { $r.icmpLoss } else { $r.lossPct })) }
$lpS      = Stats $prefLossVals

$gwS      = Stats ($data | Select-Object -ExpandProperty gwMs)
$dnsS     = Stats ($data | Select-Object -ExpandProperty dnsMs)
$icmpAvgS = Stats ($data | Select-Object -ExpandProperty icmpAvg)
$icmpMaxS = Stats ($data | Select-Object -ExpandProperty icmpMax)
$icmpMinS = Stats ($data | Select-Object -ExpandProperty icmpMin)
$icmpJitS = Stats ($data | Select-Object -ExpandProperty icmpJit)

# Compliance (% of runs that meet thresholds; ignore nulls)
function PercentOk($vals, [scriptblock]$pred){
  $have = 0; $ok = 0
  foreach($v in $vals){ if ($null -ne $v) { $have++; if (& $pred $v) { $ok++ } } }
  if ($have -eq 0) { return $null }
  [math]::Round(($ok*100.0)/$have,1)
}
$dlOK   = PercentOk ($data.dlMbps)   { param($v) $v -ge $PlanDlMbps }
$ulOK   = PercentOk ($data.ulMbps)   { param($v) $v -ge $PlanUlMbps }
$latOK  = PercentOk ($data.latMs)    { param($v) $v -le $LatencyMaxMs }
$lossOK = PercentOk $prefLossVals     { param($v) $v -le $LossMaxPct }
$gwOK   = PercentOk ($data.gwMs)     { param($v) $v -le $GatewayPingMaxMs }
$dnsOK  = PercentOk ($data.dnsMs)    { param($v) $v -le $DnsLookupMaxMs }

# Health colors per run (overall) — use preferred loss
$colors = @()
foreach($r in $data){
  $bad = $false; $warn = $false
  if ($r.dlMbps -ne $null) { if ($r.dlMbps -lt $PlanDlMbps) { if ($r.dlMbps -lt ($PlanDlMbps/2)) { $bad=$true } else { $warn=$true } } }
  if ($r.latMs  -ne $null) { if ($r.latMs  -gt $LatencyMaxMs) { if ($r.latMs -gt ($LatencyMaxMs*2)) { $bad=$true } else { $warn=$true } } }
  $lossVal = ($(if ($r.icmpLoss -ne $null) { $r.icmpLoss } else { $r.lossPct }))
  if ($lossVal -ne $null){ if ($lossVal -gt $LossMaxPct) { if ($lossVal -gt ($LossMaxPct*3)) { $bad=$true } else { $warn=$true } } }
  if ($r.gwMs -ne $null)   { if ($r.gwMs -gt $GatewayPingMaxMs*2) { $warn=$true } }
  if ($r.dnsMs -ne $null)  { if ($r.dnsMs -gt $DnsLookupMaxMs*2)  { $warn=$true } }
  if ($r.error) { $bad = $true }
  $colors += ($(if ($bad) { '#ff6371' } elseif ($warn) { '#ffd166' } else { '#06d6a0' }))
}
$heat = New-HeatStrip -Count $rows -Colors $colors

# ===== Series for charts =====
$labels = $data.ts | ForEach-Object { $_.ToString('MM-dd HH:mm') }  # not drawn, for info only
$dl     = $data.dlMbps
$ul     = $data.ulMbps
$lat    = $data.latMs
$jit    = $data.jitMs
# prefer ICMP loss for plotting
$loss   = @(); foreach($r in $data){ $loss += ($(if ($r.icmpLoss -ne $null) { $r.icmpLoss } else { $r.lossPct })) }
$gw     = $data.gwMs
$dns    = $data.dnsMs

$svgRate    = New-SvgChart -Title 'Throughput' -YUnit 'Mbps' -Threshold $PlanDlMbps -Series @(
  @{ name='Download'; values=$dl; stroke='#4FC3F7' },
  @{ name='Upload';   values=$ul; stroke='#81C784' }
)
$svgLatency = New-SvgChart -Title 'Latency & Jitter' -YUnit 'ms' -Threshold $LatencyMaxMs -Series @(
  @{ name='Latency'; values=$lat; stroke='#FDD835' },
  @{ name='Jitter';  values=$jit; stroke='#FFB74D' }
)
$svgLoss    = New-SvgChart -Title 'Packet Loss (preferred ICMP)' -YUnit '%' -Threshold $LossMaxPct -Series @(
  @{ name='Loss'; values=$loss; stroke='#EF9A9A' }
)
$svgPath    = New-SvgChart -Title 'Path Health (Gateway & DNS)' -YUnit 'ms' -Threshold $GatewayPingMaxMs -Series @(
  @{ name='Gateway ping'; values=$gw;  stroke='#90CAF9' },
  @{ name='DNS lookup';   values=$dns; stroke='#CE93D8' }
)

# Last 20 rows (most recent first)
$last20 = $data | Sort-Object ts -Descending | Select-Object -First 20
$tableRows = ''
foreach ($r in $last20) {
  $link = if ($r.url) { "<a href='" + (HtmlEscape $r.url) + "' target='_blank'>link</a>" } else { "" }
  $lossCell = ($(if ($r.icmpLoss -ne $null) { $r.icmpLoss } else { $r.lossPct }))
  $tableRows += "<tr>" +
    "<td>" + (HtmlEscape ($r.ts.ToString('yyyy-MM-dd HH:mm'))) + "</td>" +
    "<td>" + (HtmlEscape $r.dlMbps) + "</td>" +
    "<td>" + (HtmlEscape $r.ulMbps) + "</td>" +
    "<td>" + (HtmlEscape $r.latMs) + "</td>" +
    "<td>" + (HtmlEscape $r.jitMs) + "</td>" +
    "<td>" + (HtmlEscape $lossCell) + "</td>" +
    "<td>" + (HtmlEscape $r.gwMs) + "</td>" +
    "<td>" + (HtmlEscape $r.dnsMs) + "</td>" +
    "<td>" + (HtmlEscape $r.server) + "</td>" +
    "<td>" + (HtmlEscape $r.loc) + "</td>" +
    "<td>" + (HtmlEscape $r.isp) + "</td>" +
    "<td>" + (HtmlEscape $r.retry) + "</td>" +
    "<td>" + $link + "</td>" +
    "<td>" + (HtmlEscape $r.error) + "</td>" +
    # ICMP extras at the end:
    "<td>" + (HtmlEscape $r.icmpAvg) + "</td>" +
    "<td>" + (HtmlEscape $r.icmpMax) + "</td>" +
    "<td>" + (HtmlEscape $r.icmpJit) + "</td>" +
  "</tr>`n"
}

# Distincts
$serverSet = ($data | Group-Object server, serverId | ForEach-Object {
  $n=$_.Group[0].server; $i=$_.Group[0].serverId
  if ($n -and $i) { "$n (ID $i)" } elseif ($n) { "$n" } else { "" } }) -join ', '
$ispSet = ($data | Group-Object isp | ForEach-Object { $_.Name }) -join ', '

# ===== HTML =====
$now = Get-Date
$logDirEsc = $LogDir -replace '\\','\\'

$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>SecureNet &ndash; Customer Network Report</title>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<style>
 body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Inter,Arial,sans-serif;margin:24px;background:#0b0f14;color:#e6eef6}
 h1,h2{margin:0 0 8px}
 .sub{opacity:.85;margin-bottom:14px}
 .badges{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0 18px}
 .badge{border:1px solid #1c2532;background:#121923;border-radius:999px;padding:6px 10px;font-size:12px}
 .ok{border-color:#135e48;background:#0e1f1a}
 .warn{border-color:#5e5a13;background:#1f1e0e}
 .grid{display:grid;gap:18px}
 .grid-2{grid-template-columns:repeat(2,minmax(0,1fr))}
 .card{background:#121923;border:1px solid #1c2532;border-radius:14px;padding:16px;box-shadow:0 2px 8px rgba(0,0,0,.25)}
 table{width:100%;border-collapse:collapse;margin-top:10px}
 th,td{padding:8px 10px;border-bottom:1px solid #1c2532;font-size:13px}
 th{background:#0e141c;text-align:left}
 @media(max-width:1000px){.grid-2{grid-template-columns:1fr}}
</style>
</head>
<body>
  <h1>SecureNet &ndash; Customer Network Report</h1>
  <div class="sub">
    Window: last ${DaysBack} day(s) &bull; Generated: ${now} &bull; Rows: ${rows} &bull; First: ${firstTs} &bull; Last: ${lastTs}<br/>
    Servers: ${serverSet} &bull; ISP(s): ${ispSet}
  </div>

  <div class="badges">
    <div class="badge $(if($dlOK -ne $null -and $dlOK -ge 95){'ok'}else{'warn'})">DL OK: $(Fmt $dlOK)% (threshold ${PlanDlMbps} Mbps)</div>
    <div class="badge $(if($ulOK -ne $null -and $ulOK -ge 95){'ok'}else{'warn'})">UL OK: $(Fmt $ulOK)% (threshold ${PlanUlMbps} Mbps)</div>
    <div class="badge $(if($latOK -ne $null -and $latOK -ge 95){'ok'}else{'warn'})">Latency OK: $(Fmt $latOK)% (&le; ${LatencyMaxMs} ms)</div>
    <div class="badge $(if($lossOK -ne $null -and $lossOK -ge 95){'ok'}else{'warn'})">Loss OK: $(Fmt $lossOK)% (&le; ${LossMaxPct}%)</div>
    <div class="badge $(if($gwOK -ne $null -and $gwOK -ge 95){'ok'}else{'warn'})">Gateway OK: $(Fmt $gwOK)% (&le; ${GatewayPingMaxMs} ms)</div>
    <div class="badge $(if($dnsOK -ne $null -and $dnsOK -ge 95){'ok'}else{'warn'})">DNS OK: $(Fmt $dnsOK)% (&le; ${DnsLookupMaxMs} ms)</div>
  </div>

  <div class="card">
    <h2>Overall Health (newest on right)</h2>
    $heat
  </div>

  <div class="grid grid-2" style="margin-top:18px">
    <div class="card"><h2>Throughput</h2>$svgRate</div>
    <div class="card"><h2>Latency &amp; Jitter</h2>$svgLatency</div>
    <div class="card"><h2>Packet Loss</h2>$svgLoss</div>
    <div class="card"><h2>Path Health (Gateway &amp; DNS)</h2>$svgPath</div>
  </div>

  <div class="card" style="margin-top:18px">
    <h2>KPIs (last ${DaysBack} day(s))</h2>
    <table>
      <thead><tr>
        <th>Metric</th><th>Min</th><th>Avg</th><th>Max</th><th>Threshold</th>
      </tr></thead>
      <tbody>
        <tr><td>Download (Mbps)</td><td>$(Fmt $dlS.min)</td><td>$(Fmt $dlS.avg)</td><td>$(Fmt $dlS.max)</td><td>&ge; ${PlanDlMbps}</td></tr>
        <tr><td>Upload (Mbps)</td><td>$(Fmt $ulS.min)</td><td>$(Fmt $ulS.avg)</td><td>$(Fmt $ulS.max)</td><td>&ge; ${PlanUlMbps}</td></tr>
        <tr><td>Latency (ms)</td><td>$(Fmt $ltS.min)</td><td>$(Fmt $ltS.avg)</td><td>$(Fmt $ltS.max)</td><td>&le; ${LatencyMaxMs}</td></tr>
        <tr><td>Jitter (ms)</td><td>$(Fmt $jtS.min)</td><td>$(Fmt $jtS.avg)</td><td>$(Fmt $jtS.max)</td><td>(informational)</td></tr>
        <tr><td>Packet Loss (%)</td><td>$(Fmt $lpS.min)</td><td>$(Fmt $lpS.avg)</td><td>$(Fmt $lpS.max)</td><td>&le; ${LossMaxPct}</td></tr>
        <tr><td>Gateway Ping (ms)</td><td>$(Fmt $gwS.min)</td><td>$(Fmt $gwS.avg)</td><td>$(Fmt $gwS.max)</td><td>&le; ${GatewayPingMaxMs}</td></tr>
        <tr><td>DNS Lookup (ms)</td><td>$(Fmt $dnsS.min)</td><td>$(Fmt $dnsS.avg)</td><td>$(Fmt $dnsS.max)</td><td>&le; ${DnsLookupMaxMs}</td></tr>
        <tr><td>ICMP Ping (ms)</td><td>$(Fmt $icmpMinS.min)</td><td>$(Fmt $icmpAvgS.avg)</td><td>$(Fmt $icmpMaxS.max)</td><td>(PingSampler 1 Hz)</td></tr>
        <tr><td>ICMP Jitter (ms)</td><td>$(Fmt $icmpJitS.min)</td><td>$(Fmt $icmpJitS.avg)</td><td>$(Fmt $icmpJitS.max)</td><td>(abs. successive diff)</td></tr>
      </tbody>
    </table>
  </div>

  <div class="card" style="margin-top:18px">
    <h2>Last 20 Runs (most recent first)</h2>
    <table>
      <thead><tr>
        <th>Timestamp</th><th>DL</th><th>UL</th><th>Lat</th><th>Jit</th><th>Loss%</th><th>GW ms</th><th>DNS ms</th><th>Server</th><th>Location</th><th>ISP</th><th>Retry</th><th>Result</th><th>Error</th><th>ICMP avg</th><th>ICMP max</th><th>ICMP jitter</th>
      </tr></thead>
      <tbody>
        $tableRows
      </tbody>
    </table>
  </div>

  <div style="opacity:.7;margin-top:12px;font-size:12px">Source CSVs: <code>$logDirEsc</code>. Print with Ctrl+P to PDF.</div>
</body>
</html>
"@

$html | Out-File -FilePath $OutHtml -Encoding UTF8
Write-Host "Report built: $OutHtml"
