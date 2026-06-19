# ============================================================
# GTS ESM Daily Case Inflow - Auto-Generator
# Matches reference design: dark navy, green/pink/blue color tiers,
# AVG/DAY + ILT columns, summary chips in header
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_Daily_Inflow.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$CasesXlsx  = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$OutputDir  = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"

# Derive date from filename
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

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing daily inflow for $Month..." -ForegroundColor Yellow

$deKeywords = @{
    'mjay'   = 'Mritunjay Kumar Singh'
    'shub'   = 'Shubham Thakur'
    'ani'    = 'Anjali Tripathi'
    'swarna' = 'Swarna Gupta'
    'kriti'  = 'Kriti Bhalla'
    'naren'  = 'Narendran Raja'
}

function Resolve-Processor($row) {
    $procId = $row.PROCESSOR_ID
    $misc   = $row.EXT_MISCINFO.Trim().ToLower()
    $name   = $row.PROCESSOR_NAME.Trim()

    if ($procId -eq 'SAP_GERMANY_USER') {
        foreach ($key in $deKeywords.Keys) {
            if ($misc -match $key) { return @{ Name = $deKeywords[$key]; IsDE = $true } }
        }
        return @{ Name = 'Germany (Unassigned)'; IsDE = $true }
    }

    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq 'None' -or $name -eq 'SAP_GERMANY_USER') {
        return @{ Name = 'Unassigned'; IsDE = $false }
    }
    return @{ Name = $name; IsDE = $false }
}

$juneCases = $cases | Where-Object { $_.MONTH -eq $Month }
Write-Host "  June cases found: $($juneCases.Count)" -ForegroundColor Gray

# Build per-processor daily counts + ILT count
$procDays = @{}
$procIsDE = @{}
$procILT  = @{}

foreach ($row in $juneCases) {
    try { $dt = [datetime]::Parse($row.CREATED_ON_DATE) } catch { continue }
    $day = $dt.Day
    if ($day -lt 1 -or $day -gt $MaxDay) { continue }

    $resolved = Resolve-Processor $row
    $pname = $resolved.Name
    $isDE  = $resolved.IsDE

    if (-not $procDays.ContainsKey($pname)) {
        $procDays[$pname] = New-Object int[] ($MaxDay + 1)
        $procIsDE[$pname] = $isDE
        $procILT[$pname]  = 0
    }
    $procDays[$pname][$day]++

    # ILT flag: SUBJECT or EXT_MISCINFO contains 'ILT'
    if ($row.SUBJECT -match 'ILT' -or $row.EXT_MISCINFO -match 'ILT') { $procILT[$pname]++ }
}

# Sort: by total desc
$allRows = @()
foreach ($pname in $procDays.Keys) {
    $total = 0
    for ($d = 1; $d -le $MaxDay; $d++) { $total += $procDays[$pname][$d] }
    $allRows += [PSCustomObject]@{ Name=$pname; Total=$total; IsDE=$procIsDE[$pname]; ILT=$procILT[$pname] }
}
$allRows = $allRows | Sort-Object Total -Descending

$totalCases = ($allRows | Measure-Object -Property Total -Sum).Sum
$agentCount = $allRows.Count
Write-Host "  Processors found: $agentCount, Total: $totalCases" -ForegroundColor Gray

# Pre-compute top 3 / bottom 5 per day (non-zero only)
$dayTop3    = @{}
$dayBottom5 = @{}
for ($d = 1; $d -le $MaxDay; $d++) {
    $dayVals = $allRows | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; Val=$procDays[$_.Name][$d] }
    } | Where-Object { $_.Val -gt 0 } | Sort-Object Val -Descending
    $dayTop3[$d]    = @(($dayVals | Select-Object -First 3).Name)
    $dayBottom5[$d] = @(($dayVals | Select-Object -Last 5).Name)
}

# Day names - June 1 2026 = Monday
$dayNames = @('Mon','Tue','Wed','Thu','Fri','Sat','Sun')

# Build header columns
$thDays = ""
for ($d = 1; $d -le $MaxDay; $d++) {
    $dayName = $dayNames[($d - 1) % 7]
    $isWeekend = ($dayName -eq 'Sat' -or $dayName -eq 'Sun')
    $wkStyle = if ($isWeekend) { ' style="opacity:0.6;"' } else { '' }
    $thDays += "<th$wkStyle><div class='th-day'>$dayName</div><div class='th-date'>Jun $d</div></th>"
}

# Build tbody
$tbodyHtml = ""
foreach ($r in $allRows) {
    $avg = if ($MaxDay -gt 0) { [math]::Round($r.Total / $MaxDay, 1) } else { 0 }
    $iltVal = if ($r.ILT -gt 0) { $r.ILT } else { "" }
    $nameStyle = if ($r.IsDE) { ' style="color:#fbbf24;"' } else { '' }

    $tbodyHtml += "  <tr>`n"
    $tbodyHtml += "    <td class='name-col'$nameStyle>$($r.Name)</td>`n"

    for ($d = 1; $d -le $MaxDay; $d++) {
        $v = $procDays[$r.Name][$d]
        $isTop    = $dayTop3[$d] -contains $r.Name
        $isBottom = $dayBottom5[$d] -contains $r.Name
        $dayName  = $dayNames[($d - 1) % 7]
        $isWeekend = ($dayName -eq 'Sat' -or $dayName -eq 'Sun')

        if ($v -eq 0) {
            $wkOpacity = if ($isWeekend) { ' style="opacity:0.4;"' } else { '' }
            $tbodyHtml += "    <td class='cell-none'$wkOpacity>--</td>`n"
        } elseif ($isTop) {
            $tbodyHtml += "    <td class='cell-top'>$v</td>`n"
        } elseif ($isBottom) {
            $tbodyHtml += "    <td class='cell-bot'>$v</td>`n"
        } else {
            $wkClass = if ($isWeekend) { ' cell-wk' } else { '' }
            $tbodyHtml += "    <td class='cell-mid$wkClass'>$v</td>`n"
        }
    }

    $tbodyHtml += "    <td class='col-total'>$($r.Total)</td>`n"
    $tbodyHtml += "    <td class='col-avg'>$avg</td>`n"
    $tbodyHtml += "    <td class='col-ilt'>$(if($iltVal){"$iltVal"}else{'--'})</td>`n"
    $tbodyHtml += "  </tr>`n"
}

# Build tfoot
$tfootHtml = "  <tr>`n    <td class='name-col'>DAILY TOTAL</td>`n"
$grandTotal = 0
$colTotals  = @()
for ($d = 1; $d -le $MaxDay; $d++) {
    $daySum = 0
    foreach ($pname in $procDays.Keys) { $daySum += $procDays[$pname][$d] }
    $colTotals += $daySum
    $grandTotal += $daySum
    $tfootHtml += "    <td class='foot-cell'>$daySum</td>`n"
}
$grandAvg = [math]::Round($grandTotal / $MaxDay, 1)
$grandILT = ($allRows | Measure-Object -Property ILT -Sum).Sum
$tfootHtml += "    <td class='col-total foot-total'>$grandTotal</td>`n"
$tfootHtml += "    <td class='col-avg foot-avg'>$grandAvg</td>`n"
$tfootHtml += "    <td class='col-ilt foot-ilt'>$(if($grandILT -gt 0){$grandILT}else{'--'})</td>`n"
$tfootHtml += "  </tr>`n"

$outputFile = "$OutputDir\GTS_Daily_Inflow_June.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>GTS June Daily Inflow</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a1a2e; padding: 0; }

    /* ---- PAGE HEADER ---- */
    .page-header {
      background: #1e3a5f;
      padding: 18px 24px 14px;
    }
    .page-header h1 { font-size: 22px; font-weight: 700; color: #fff; margin-bottom: 10px; }
    .chips { display: flex; gap: 10px; flex-wrap: wrap; }
    .chip {
      display: flex; align-items: center; gap: 6px;
      background: rgba(255,255,255,0.15); border-radius: 20px;
      padding: 4px 12px; font-size: 12px; color: #e0f0ff;
    }
    .chip .dot { width: 8px; height: 8px; border-radius: 50%; background: #7dd3fc; flex-shrink: 0; }

    /* ---- SECTION LABEL ---- */
    .section-bar {
      background: #fff; padding: 8px 24px;
      font-size: 12px; color: #6b7280;
      border-bottom: 1px solid #e5e7eb;
      display: flex; align-items: center; gap: 12px;
    }
    .section-bar strong { color: #1e3a5f; }
    .tag-chip {
      background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 4px;
      padding: 2px 8px; font-size: 11px; color: #2563eb;
    }

    /* ---- TABLE WRAP ---- */
    .table-wrap { overflow-x: auto; padding: 0 0 16px; background: #fff; }

    table { width: 100%; border-collapse: collapse; font-size: 12px; white-space: nowrap; }

    /* ---- THEAD ---- */
    thead tr { background: #1e3a5f; }
    thead th {
      padding: 8px 6px 6px;
      text-align: center;
      font-weight: 600;
      color: #bfdbfe;
      border-bottom: 2px solid #2e5480;
      border-right: 1px solid #2e5480;
      min-width: 58px;
    }
    thead th.name-col {
      text-align: left; padding-left: 20px;
      min-width: 200px; color: #fff;
    }
    thead th.col-total { background: #162d4a; color: #7dd3fc; min-width: 70px; }
    thead th.col-avg   { background: #162d4a; color: #c4b5fd; min-width: 65px; }
    thead th.col-ilt   { background: #4c1d6a; color: #e9d5ff; min-width: 50px; }
    .th-day  { font-size: 10px; font-weight: 400; color: #93c5fd; margin-bottom: 2px; }
    .th-date { font-size: 12px; font-weight: 700; color: #fff; }

    /* ---- TBODY ---- */
    tbody tr { border-bottom: 1px solid #f3f4f6; background: #fff; }
    tbody tr:nth-child(even) { background: #f9fafb; }
    tbody tr:hover { background: #eff6ff !important; }

    td {
      padding: 7px 6px;
      text-align: center;
      border-right: 1px solid #f3f4f6;
      color: #374151;
    }
    td.name-col {
      text-align: left; padding-left: 20px;
      font-weight: 600; color: #1e3a5f;
      background: inherit;
      position: sticky; left: 0; z-index: 1;
      border-right: 2px solid #e5e7eb;
    }

    /* color cells */
    td.cell-top  { background: #166534; color: #fff; font-weight: 700; }
    td.cell-bot  { background: #fca5a5; color: #7f1d1d; font-weight: 600; }
    td.cell-mid  { background: #bfdbfe; color: #1e3a5f; font-weight: 500; }
    td.cell-none { color: #d1d5db; font-size: 11px; }

    /* totals cols */
    td.col-total { background: #eff6ff; color: #1d4ed8; font-weight: 700; font-size: 13px; border-left: 2px solid #bfdbfe; }
    td.col-avg   { background: #f5f3ff; color: #6d28d9; font-weight: 600; }
    td.col-ilt   { background: #faf5ff; color: #7e22ce; font-weight: 600; }

    /* ---- TFOOT ---- */
    tfoot tr { background: #1e3a5f; border-top: 2px solid #2e5480; }
    tfoot td {
      padding: 9px 6px;
      text-align: center;
      font-weight: 700; font-size: 12px;
      color: #e2e8f0;
      border-right: 1px solid #2e5480;
    }
    tfoot td.name-col {
      text-align: left; padding-left: 20px;
      font-size: 13px; letter-spacing: 0.5px;
      color: #fff; background: #1e3a5f;
      text-transform: uppercase;
    }
    tfoot td.foot-total { background: #162d4a; color: #7dd3fc; font-size: 14px; }
    tfoot td.foot-avg   { background: #162d4a; color: #c4b5fd; }
    tfoot td.foot-ilt   { background: #3b0764; color: #e9d5ff; }
    tfoot .foot-cell    { color: #cbd5e1; }

    /* ---- LEGEND ---- */
    .legend {
      display: flex; gap: 20px; align-items: center;
      padding: 10px 24px; font-size: 11px; color: #6b7280;
      border-top: 1px solid #e5e7eb; background: #fff;
    }
    .legend-item { display: flex; align-items: center; gap: 6px; }
    .legend-swatch { width: 14px; height: 14px; border-radius: 3px; }
  </style>
</head>
<body>

<div class="page-header">
  <h1>GTS June Daily Inflow</h1>
  <div class="chips">
    <div class="chip"><div class="dot"></div> June 1-$MaxDay, 2026</div>
    <div class="chip"><div class="dot" style="background:#7dd3fc"></div> $($totalCases.ToString('N0')) total cases</div>
    <div class="chip"><div class="dot" style="background:#c4b5fd"></div> $agentCount agents</div>
    <div class="chip"><div class="dot" style="background:#6ee7b7"></div> $MaxDay days</div>
    <div class="chip"><div class="dot" style="background:#fde68a"></div> Source: $(Split-Path $CasesXlsx -Leaf)</div>
  </div>
</div>

<div class="section-bar">
  <strong>Agent Daily Inflow</strong>
  <span class="tag-chip">rows = agents</span>
  <span class="tag-chip">columns = dates</span>
</div>

<div class="table-wrap">
<table>
  <thead>
    <tr>
      <th class="name-col">Agent</th>
      $thDays
      <th class="col-total">TOTAL</th>
      <th class="col-avg">AVG<br>/DAY</th>
      <th class="col-ilt">ILT</th>
    </tr>
  </thead>
  <tbody>
$tbodyHtml  </tbody>
  <tfoot>
$tfootHtml  </tfoot>
</table>
</div>

<div class="legend">
  <div class="legend-item"><div class="legend-swatch" style="background:#166534"></div> Top 3 that day</div>
  <div class="legend-item"><div class="legend-swatch" style="background:#fca5a5"></div> Bottom 5 that day</div>
  <div class="legend-item"><div class="legend-swatch" style="background:#bfdbfe"></div> Mid</div>
  <div class="legend-item"><div class="legend-swatch" style="background:#f3f4f6;border:1px solid #d1d5db"></div> None</div>
</div>

</body>
</html>
"@

$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
