# ============================================================
# GTS ESM Daily Case Inflow + Closures - Auto-Generator
# Layout: W1 (Jun 1-7) | W2 (Jun 8-14) | daily from Jun 15+
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_Daily_Inflow.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$CasesXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx" | Where-Object { $_.Name -notmatch 'with closed' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$OutputDir = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"

$baseName   = [System.IO.Path]::GetFileNameWithoutExtension($CasesXlsx)
$datePart   = $baseName -replace '^GTS ESM Cases_', ''
$parsedDate = $null
foreach ($fmt in @('MMMMd','MMMMdd','MMMd','MMMdd')) {
    try { $parsedDate = [datetime]::ParseExact($datePart, $fmt, [System.Globalization.CultureInfo]::InvariantCulture); break } catch {}
}
if (-not $parsedDate) { $parsedDate = [datetime]::Today }
$parsedDate = $parsedDate.AddYears([datetime]::Today.Year - $parsedDate.Year)
$ReportDate = $parsedDate.ToString("MMMM d, yyyy")
$Month      = $parsedDate.ToString("MMMM")
$MaxDay     = $parsedDate.Day

# Week boundaries
$W1Start = 1;  $W1End = 7
$W2Start = 8;  $W2End = 14
$W3Start = 15  # current week — daily from here to MaxDay
# -- END CONFIG -----------------------------------------------

$CasesCsv = "$env:TEMP\gts_cases_inflow.csv"

function Export-XlsxToCsv($xlsxPath, $csvPath) {
    Write-Host "  Exporting: $xlsxPath" -ForegroundColor Cyan
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($xlsxPath)
    $wb.SaveAs($csvPath, 6)
    $wb.Close($false); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

Write-Host "Exporting data..." -ForegroundColor Yellow
Export-XlsxToCsv $CasesXlsx $CasesCsv

$peekContent = Get-Content $CasesCsv -TotalCount 3
$HasClosures = ($peekContent -join ' ') -match 'COMPLETED_ON|CLOSED_ON'
Write-Host "  Closure columns detected: $HasClosures" -ForegroundColor Gray

if ($HasClosures) {
    $cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
              'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
              'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
              'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','COMPLETED_ON','CLOSED_ON','IRT_HOURS')
} else {
    $cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
              'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
              'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
              'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
}
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing inflow/closures for $Month..." -ForegroundColor Yellow

$deKeywords = @{
    'mjay'='Mritunjay Kumar Singh'; 'mritunjay'='Mritunjay Kumar Singh'
    'shub'='Shubham Thakur';        'shubham'='Shubham Thakur'
    'ani'='Anjali Tripathi';        'anjali'='Anjali Tripathi'
    'swarna'='Swarna Gupta';        'kriti'='Kriti Bhalla'
    'naren'='Narendran Raja';       'narendran'='Narendran Raja'
}

function Resolve-Processor($row) {
    $procId = $row.PROCESSOR_ID
    $misc   = $row.EXT_MISCINFO.Trim().ToLower()
    $name   = $row.PROCESSOR_NAME.Trim()
    if ($procId -eq 'SAP_GERMANY_USER') {
        foreach ($key in $deKeywords.Keys) {
            if ($misc -match $key) { return @{ Name=$deKeywords[$key]; IsDE=$true } }
        }
        return @{ Name='Germany (Unassigned)'; IsDE=$true }
    }
    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq 'None' -or $name -eq 'SAP_GERMANY_USER') {
        return @{ Name='Unassigned'; IsDE=$false }
    }
    return @{ Name=$name; IsDE=$false }
}

# Current week daily columns
$dailyCols = @()
for ($d = $W3Start; $d -le $MaxDay; $d++) { $dailyCols += $d }

# Storage: inflow and closure per agent per day
$procInflow  = @{}  # [agent][day]
$procClose   = @{}  # [agent][day]
$procIsDE    = @{}
$procILT     = @{}

$monthCases = $cases | Where-Object { $_.MONTH -eq $Month }
Write-Host "  $Month cases found: $($monthCases.Count)" -ForegroundColor Gray

foreach ($row in $monthCases) {
    try { $dt = [datetime]::Parse($row.CREATED_ON_DATE) } catch { continue }
    $day = $dt.Day
    if ($day -lt 1 -or $day -gt $MaxDay) { continue }
    $r = Resolve-Processor $row
    $p = $r.Name
    if (-not $procInflow.ContainsKey($p)) {
        $procInflow[$p] = New-Object int[] ($MaxDay + 1)
        $procClose[$p]  = New-Object int[] ($MaxDay + 1)
        $procIsDE[$p]   = $r.IsDE
        $procILT[$p]    = 0
    }
    $procInflow[$p][$day]++
    if ($row.SUBJECT -match 'ILT' -or $row.EXT_MISCINFO -match 'ILT') { $procILT[$p]++ }
}

if ($HasClosures) {
    foreach ($row in $cases) {
        $doneDate = if ($row.COMPLETED_ON -ne '') { $row.COMPLETED_ON } else { $row.CLOSED_ON }
        if ([string]::IsNullOrWhiteSpace($doneDate)) { continue }
        try { $dt = [datetime]::Parse($doneDate) } catch { continue }
        if ($dt.ToString('MMMM') -ne $Month) { continue }
        $day = $dt.Day
        if ($day -lt 1 -or $day -gt $MaxDay) { continue }
        $r = Resolve-Processor $row
        $p = $r.Name
        if (-not $procClose.ContainsKey($p)) { $procClose[$p] = New-Object int[] ($MaxDay + 1) }
        $procClose[$p][$day]++
        if (-not $procInflow.ContainsKey($p)) {
            $procInflow[$p] = New-Object int[] ($MaxDay + 1)
            $procIsDE[$p]   = $r.IsDE
            $procILT[$p]    = 0
        }
    }
}

# Aggregate weekly totals
function WeekSum($arr, $s, $e) {
    $t = 0; for ($d = $s; $d -le $e; $d++) { $t += $arr[$d] }; return $t
}

# Sort agents by total inflow
$allRows = @()
foreach ($p in $procInflow.Keys) {
    $tot = 0; for ($d=1;$d-le$MaxDay;$d++){$tot+=$procInflow[$p][$d]}
    $allRows += [PSCustomObject]@{ Name=$p; Total=$tot; IsDE=$procIsDE[$p]; ILT=$procILT[$p] }
}
$allRows = $allRows | Sort-Object Total -Descending

$totalInflow = ($allRows | Measure-Object Total -Sum).Sum
$totalClose  = 0
if ($HasClosures) {
    foreach ($p in $procClose.Keys) { for($d=1;$d-le$MaxDay;$d++){$totalClose+=$procClose[$p][$d]} }
}
$agentCount = $allRows.Count

Write-Host "  Agents: $agentCount | Total inflow: $totalInflow | Total closures: $totalClose" -ForegroundColor Gray

# Day name helper (June 1 2026 = Monday)
$dayNames  = @('Mon','Tue','Wed','Thu','Fri','Sat','Sun')
function DayName($d) { $dayNames[($d-1)%7] }
function IsWeekend($d) { $n = DayName $d; $n -eq 'Sat' -or $n -eq 'Sun' }

# Build table header
$thCols = ""
$thCols += "<th class='wk-col w1'>W1<br><span class='wk-sub'>Jun 1-7</span></th>"
$thCols += "<th class='wk-col w2'>W2<br><span class='wk-sub'>Jun 8-14</span></th>"
foreach ($d in $dailyCols) {
    $dn = DayName $d
    $wkStyle = if (IsWeekend $d) { ' style="opacity:0.65"' } else { '' }
    $thCols += "<th class='day-col'$wkStyle><div class='th-day'>$dn</div><div class='th-date'>$d</div></th>"
}
$thCols += "<th class='col-total'>TOTAL</th><th class='col-avg'>AVG<br>/DAY</th><th class='col-ilt'>ILT</th><th class='col-rate'>RATE%</th>"

# Build tbody
$tbodyHtml = ""
foreach ($r in $allRows) {
    $nameStyle = if ($r.IsDE) { ' style="color:#fbbf24"' } else { '' }
    $iw1 = WeekSum $procInflow[$r.Name] $W1Start $W1End
    $iw2 = WeekSum $procInflow[$r.Name] $W2Start $W2End
    $cw1 = if ($HasClosures) { WeekSum $procClose[$r.Name] $W1Start $W1End } else { 0 }
    $cw2 = if ($HasClosures) { WeekSum $procClose[$r.Name] $W2Start $W2End } else { 0 }
    $itot = 0; for($d=1;$d-le$MaxDay;$d++){$itot+=$procInflow[$r.Name][$d]}
    $ctot = 0; if($HasClosures){for($d=1;$d-le$MaxDay;$d++){$ctot+=$procClose[$r.Name][$d]}}
    $iavg = [math]::Round($itot/$MaxDay,1)
    $cavg = [math]::Round($ctot/$MaxDay,1)

    $rate = if ($itot -gt 0) { [math]::Round($ctot / $itot * 100) } else { 0 }
    $rateClass = if ($rate -ge 100) { 'rate-good' } elseif ($rate -ge 80) { 'rate-warn' } else { 'rate-bad' }

    # INFLOW ROW
    $tbodyHtml += "<tr class='row-in'>"
    $tbodyHtml += "<td class='name-col'$nameStyle><span class='lbl lbl-in'>IN</span> $($r.Name)</td>"
    $tbodyHtml += "<td class='wk-cell'>$iw1</td><td class='wk-cell'>$iw2</td>"
    foreach ($d in $dailyCols) {
        $v = $procInflow[$r.Name][$d]
        $wk = if (IsWeekend $d) { ' wk' } else { '' }
        if ($v -eq 0) { $tbodyHtml += "<td class='z$wk'>--</td>" }
        else          { $tbodyHtml += "<td class='in-cell$wk'>$v</td>" }
    }
    $tbodyHtml += "<td class='col-total'>$itot</td><td class='col-avg'>$iavg</td><td class='col-ilt'>$(if($r.ILT -gt 0){$r.ILT}else{'--'})</td>"
    $tbodyHtml += "<td class='col-rate $rateClass' rowspan='2'>$rate%</td>"
    $tbodyHtml += "</tr>`n"

    # CLOSURE ROW
    if ($HasClosures) {
        $tbodyHtml += "<tr class='row-cls'>"
        $tbodyHtml += "<td class='name-col cls-name'><span class='lbl lbl-cls'>CLS</span></td>"
        $tbodyHtml += "<td class='wk-cell cls-wk'>$cw1</td><td class='wk-cell cls-wk'>$cw2</td>"
        foreach ($d in $dailyCols) {
            $v = $procClose[$r.Name][$d]
            $wk = if (IsWeekend $d) { ' wk' } else { '' }
            if ($v -eq 0) { $tbodyHtml += "<td class='z cls-z$wk'>--</td>" }
            else          { $tbodyHtml += "<td class='cls-cell$wk'>$v</td>" }
        }
        $tbodyHtml += "<td class='col-total cls-total'>$ctot</td><td class='col-avg cls-avg'>$cavg</td><td class='col-ilt'>--</td>"
        $tbodyHtml += "</tr>`n"
        $tbodyHtml += "<tr class='spacer'><td colspan='999'></td></tr>`n"
    }
}

# TFOOT
$tfootHtml = ""
# Inflow foot
$tfootHtml += "<tr class='foot-in'><td class='name-col'>INFLOW TOTAL</td>"
$fw1i=0; for($d=$W1Start;$d-le$W1End;$d++){foreach($p in $procInflow.Keys){$fw1i+=$procInflow[$p][$d]}}
$fw2i=0; for($d=$W2Start;$d-le$W2End;$d++){foreach($p in $procInflow.Keys){$fw2i+=$procInflow[$p][$d]}}
$tfootHtml += "<td class='wk-cell'>$fw1i</td><td class='wk-cell'>$fw2i</td>"
$gTotal=0
foreach ($d in $dailyCols) {
    $ds=0; foreach($p in $procInflow.Keys){$ds+=$procInflow[$p][$d]}; $gTotal+=$ds
    $tfootHtml += "<td class='foot-cell'>$ds</td>"
}
$gTotal += $fw1i + $fw2i
$gAvg = [math]::Round($totalInflow/$MaxDay,1)
$gILT = ($allRows | Measure-Object ILT -Sum).Sum
$gRate = if ($totalInflow -gt 0) { [math]::Round($totalClose / $totalInflow * 100) } else { 0 }
$gRateClass = if ($gRate -ge 100) { 'rate-good' } elseif ($gRate -ge 80) { 'rate-warn' } else { 'rate-bad' }
$tfootHtml += "<td class='col-total foot-total'>$totalInflow</td><td class='col-avg foot-avg'>$gAvg</td><td class='col-ilt foot-ilt'>$(if($gILT -gt 0){$gILT}else{'--'})</td><td class='col-rate $gRateClass'>$gRate%</td></tr>`n"

if ($HasClosures) {
    $tfootHtml += "<tr class='foot-cls'><td class='name-col' style='background:#0a3d2b;color:#6ee7b7'>CLOSURE TOTAL</td>"
    $fw1c=0; for($d=$W1Start;$d-le$W1End;$d++){foreach($p in $procClose.Keys){$fw1c+=$procClose[$p][$d]}}
    $fw2c=0; for($d=$W2Start;$d-le$W2End;$d++){foreach($p in $procClose.Keys){$fw2c+=$procClose[$p][$d]}}
    $tfootHtml += "<td class='wk-cell' style='background:#0a3d2b;color:#6ee7b7'>$fw1c</td><td class='wk-cell' style='background:#0a3d2b;color:#6ee7b7'>$fw2c</td>"
    foreach ($d in $dailyCols) {
        $ds=0; foreach($p in $procClose.Keys){$ds+=$procClose[$p][$d]}
        $tfootHtml += "<td class='foot-cell' style='background:#0a3d2b;color:#6ee7b7'>$ds</td>"
    }
    $cAvg = [math]::Round($totalClose/$MaxDay,1)
    $tfootHtml += "<td class='col-total' style='background:#062818;color:#34d399;font-weight:700'>$totalClose</td><td class='col-avg' style='background:#062818;color:#34d399'>$cAvg</td><td class='col-ilt foot-ilt'>--</td><td></td></tr>`n"
}

$outputFile = "$OutputDir\GTS_Daily_Inflow_June.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>GTS June Inflow & Closures</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1a1a2e}

    .page-header{background:#1e3a5f;padding:18px 24px 14px}
    .page-header h1{font-size:22px;font-weight:700;color:#fff;margin-bottom:10px}
    .chips{display:flex;gap:10px;flex-wrap:wrap}
    .chip{display:flex;align-items:center;gap:6px;background:rgba(255,255,255,0.15);border-radius:20px;padding:4px 12px;font-size:12px;color:#e0f0ff}
    .chip .dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}

    .section-bar{background:#fff;padding:8px 24px;font-size:12px;color:#6b7280;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
    .section-bar strong{color:#1e3a5f}
    .tag{background:#eff6ff;border:1px solid #bfdbfe;border-radius:4px;padding:2px 8px;font-size:11px;color:#2563eb}
    .tag.g{background:#f0fdf4;border-color:#bbf7d0;color:#16a34a}

    .table-wrap{overflow-x:auto;padding-bottom:16px;background:#fff}
    table{width:100%;border-collapse:collapse;font-size:12px;white-space:nowrap}

    /* THEAD */
    thead tr{background:#1e3a5f}
    thead th{padding:8px 8px 6px;text-align:center;font-weight:600;color:#bfdbfe;border-bottom:2px solid #2e5480;border-right:1px solid #2e5480}
    thead th.name-col{text-align:left;padding-left:16px;min-width:200px;color:#fff}
    thead th.wk-col{min-width:70px;font-size:12px}
    thead th.w1{background:#1a4a7a;color:#93c5fd}
    thead th.w2{background:#1a4a7a;color:#93c5fd}
    .wk-sub{font-size:10px;font-weight:400;color:#7dd3fc;display:block}
    thead th.day-col{min-width:50px}
    .th-day{font-size:10px;font-weight:400;color:#93c5fd;margin-bottom:2px}
    .th-date{font-size:12px;font-weight:700;color:#fff}
    thead th.col-total{background:#162d4a;color:#7dd3fc;min-width:65px}
    thead th.col-avg{background:#162d4a;color:#c4b5fd;min-width:58px}
    thead th.col-ilt{background:#4c1d6a;color:#e9d5ff;min-width:45px}

    /* ROWS */
    tr.row-in{border-top:2px solid #dde3ea;background:#fff}
    tr.row-in:hover,tr.row-cls:hover{background:#eff6ff!important}
    tr.row-cls{background:#f0fdf4}
    tr.spacer td{height:5px;background:#e5e7eb;padding:0}

    td{padding:6px 7px;text-align:center;border-right:1px solid #f3f4f6;color:#374151;font-size:12px}
    td.name-col{text-align:left;padding-left:12px;font-weight:600;color:#1e3a5f;background:inherit;position:sticky;left:0;z-index:1;border-right:2px solid #e5e7eb}
    td.cls-name{color:#059669;font-size:11px;background:#f0fdf4}

    .lbl{display:inline-block;font-size:9px;font-weight:700;border-radius:3px;padding:1px 4px;margin-right:4px}
    .lbl-in{background:#1e3a5f;color:#93c5fd}
    .lbl-cls{background:#065f46;color:#6ee7b7}

    /* WEEKLY COLS */
    td.wk-cell{background:#e8f0fe;color:#1d4ed8;font-weight:700;font-size:13px;border-right:2px solid #bfdbfe}
    td.cls-wk{background:#dcfce7;color:#166534;font-weight:700;border-right:2px solid #bbf7d0}

    /* INFLOW DAILY */
    td.in-cell{background:#bfdbfe;color:#1e3a5f;font-weight:500}
    td.in-cell.wk{opacity:0.65}
    td.z{color:#d1d5db;font-size:11px}
    td.z.wk{opacity:0.5}

    /* CLOSURE DAILY */
    td.cls-cell{background:#bbf7d0;color:#064e3b;font-weight:500}
    td.cls-cell.wk{opacity:0.65}
    td.cls-z{color:#a7f3d0;font-size:11px;background:#f0fdf4}

    /* TOTALS */
    td.col-total{background:#eff6ff;color:#1d4ed8;font-weight:700;font-size:13px;border-left:2px solid #bfdbfe}
    td.col-avg{background:#f5f3ff;color:#6d28d9;font-weight:600}
    td.col-ilt{background:#faf5ff;color:#7e22ce;font-weight:600}
    td.cls-total{background:#ecfdf5;color:#065f46;font-weight:700;font-size:13px;border-left:2px solid #6ee7b7}
    td.cls-avg{background:#ecfdf5;color:#059669;font-weight:600}

    /* TFOOT */
    tfoot tr.foot-in{background:#1e3a5f;border-top:2px solid #2e5480}
    tfoot tr.foot-cls{border-top:1px solid #065f46}
    tfoot td{padding:9px 7px;text-align:center;font-weight:700;font-size:12px;color:#e2e8f0;border-right:1px solid #2e5480}
    tfoot td.name-col{text-align:left;padding-left:16px;font-size:12px;text-transform:uppercase;letter-spacing:0.4px;background:#1e3a5f;color:#fff}
    tfoot .foot-cell{color:#cbd5e1}
    tfoot .foot-total{background:#162d4a;color:#7dd3fc;font-size:14px}
    tfoot .foot-avg{background:#162d4a;color:#c4b5fd}
    tfoot .foot-ilt{background:#3b0764;color:#e9d5ff}

    .legend{display:flex;gap:18px;align-items:center;padding:10px 24px;font-size:11px;color:#6b7280;border-top:1px solid #e5e7eb;background:#fff;flex-wrap:wrap}
    .li{display:flex;align-items:center;gap:5px}
    .sw{width:14px;height:14px;border-radius:3px}
    .div{width:1px;height:18px;background:#e5e7eb}

    thead th.col-rate{background:#1a3a1a;color:#6ee7b7;min-width:65px}
    td.col-rate{font-weight:700;font-size:13px;border-left:2px solid #e5e7eb;vertical-align:middle;text-align:center}
    td.rate-good{background:#166534;color:#fff}
    td.rate-warn{background:#854d0e;color:#fff}
    td.rate-bad{background:#7f1d1d;color:#fff}

    /* TOGGLE BUTTON */
    .toggle-bar{background:#fff;padding:8px 24px;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;gap:12px}
    .btn-toggle{display:inline-flex;align-items:center;gap:8px;background:#1e3a5f;color:#fff;border:none;border-radius:6px;padding:7px 16px;font-size:12px;font-weight:600;cursor:pointer;transition:background 0.2s}
    .btn-toggle:hover{background:#2e5480}
    .btn-toggle .dot{width:10px;height:10px;border-radius:50%;background:#6ee7b7}
    .toggle-hint{font-size:11px;color:#6b7280}
    tr.row-cls{transition:opacity 0.15s}
    tr.row-cls.hidden{display:none}
  </style>
</head>
<body>

<div class="page-header">
  <h1>GTS June Daily Inflow &amp; Closures</h1>
  <div class="chips">
    <div class="chip"><div class="dot" style="background:#7dd3fc"></div>June 1&ndash;$MaxDay, 2026</div>
    <div class="chip"><div class="dot" style="background:#7dd3fc"></div>$($totalInflow.ToString('N0')) inflow</div>
    <div class="chip"><div class="dot" style="background:#6ee7b7"></div>$($totalClose.ToString('N0')) closures</div>
    <div class="chip"><div class="dot" style="background:#c4b5fd"></div>$agentCount agents</div>
    <div class="chip"><div class="dot" style="background:#fde68a"></div>$(Split-Path $CasesXlsx -Leaf)</div>
  </div>
</div>

<div class="section-bar">
  <strong>Agent Inflow &amp; Closures</strong>
  <span class="tag">W1 / W2 = weekly totals</span>
  <span class="tag">Jun $W3Start+ = daily</span>
  <span class="tag g">CLS = closed or completed that day</span>
</div>

<div class="toggle-bar">
  <button class="btn-toggle" id="toggleBtn" onclick="toggleCLS()">
    <span class="dot"></span> Hide Closures
  </button>
  <span class="toggle-hint">Toggle to show/hide the CLS rows</span>
</div>

<div class="table-wrap">
<table>
  <thead>
    <tr>
      <th class="name-col">Agent</th>
      $thCols
    </tr>
  </thead>
  <tbody>
$tbodyHtml  </tbody>
  <tfoot>
$tfootHtml  </tfoot>
</table>
</div>

<div class="legend">
  <strong style="color:#1e3a5f">INFLOW:</strong>
  <div class="li"><div class="sw" style="background:#e8f0fe;border:1px solid #bfdbfe"></div>Weekly total</div>
  <div class="li"><div class="sw" style="background:#bfdbfe"></div>Daily</div>
  <div class="div"></div>
  <strong style="color:#065f46">CLOSURES:</strong>
  <div class="li"><div class="sw" style="background:#dcfce7;border:1px solid #bbf7d0"></div>Weekly total</div>
  <div class="li"><div class="sw" style="background:#bbf7d0"></div>Daily</div>
</div>

</body>
<script>
  var clsVisible = true;
  function toggleCLS() {
    clsVisible = !clsVisible;
    var rows = document.querySelectorAll('tr.row-cls, tr.spacer');
    rows.forEach(function(r){ r.classList.toggle('hidden', !clsVisible); });
    var btn = document.getElementById('toggleBtn');
    btn.innerHTML = clsVisible
      ? '<span class="dot"></span> Hide Closures'
      : '<span class="dot" style="background:#fbbf24"></span> Show Closures';
    btn.style.background = clsVisible ? '#1e3a5f' : '#065f46';
  }
</script>
</html>
"@

$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
