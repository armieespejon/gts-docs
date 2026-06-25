# ============================================================
# GTS Backlog Busters - Midweek Progress Report Generator
# Scoring: Closure Rate pts + Aging Score pts + CSAT Bonus
# PH Team only (13 agents)
# ============================================================

$CasesXlsx   = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx" | Where-Object { $_.Name -notmatch 'with closed' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$SurveysXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Surveys_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$OutputDir   = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"

# Week 1: Jun 22-28
$WeekStart = [datetime]'2026-06-22'
$WeekEnd   = [datetime]'2026-06-28'
$Today     = [datetime]::Today

Write-Host "Cases:   $CasesXlsx" -ForegroundColor Cyan
Write-Host "Surveys: $SurveysXlsx" -ForegroundColor Cyan

# PH agents list
$PHAgents = @(
    'Nobela, Monica Frances','Quinto, Marviely','Chua, Sam','Santos, Ana Lizelle',
    'Mendoza, Edrian','Tulagan, Nicole','Valdez, Lemuel Von','Cruz, Ann',
    'Sinag, Rendel Edison','Evidente, Marie Cristine','Buenafe, Paul',
    'Cada, Saldy','Paulino, Justin'
)

# -- EXPORT CASES CSV --
$CasesCsv = "$env:TEMP\gts_bb_cases.csv"
function Export-XlsxToCsv($xlsxPath, $csvPath) {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($xlsxPath)
    $wb.SaveAs($csvPath, 6)
    $wb.Close($false); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

Write-Host "Exporting cases..." -ForegroundColor Yellow
Export-XlsxToCsv $CasesXlsx $CasesCsv

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','COMPLETED_ON','CLOSED_ON','IRT_HOURS')
$allCases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

# -- EXPORT SURVEYS --
Write-Host "Exporting surveys..." -ForegroundColor Yellow
$SurveysCsv = "$env:TEMP\gts_bb_surveys.csv"
Export-XlsxToCsv $SurveysXlsx $SurveysCsv
$sHdr = @('DISPLAYID','CASETYPE','CASETYPEDESC','MONTH_DESC','EXT_COURSECODE','EXT_MISCINFO','SUBJECT','PRIORITY','SOURCE','SERVICETEAMNAME','ACCOUNTNAME','COMPLETED_ON_TIMESTAMP','RATING','COMMENT1','COMMENT2','RECORDEDDATE','PROCESSOR_ID','PROCESSOR_NAME','CAT1','CAT2','DOMAIN','FORWARDED_TO','COL23','COL24')
$allSurveys = Import-Csv $SurveysCsv -Header $sHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

# -- SCORING --

# 1. CLOSURE COUNT: cases closed this week (WeekStart to Today)
$weekInflow  = @{}
$weekClosure = @{}
foreach ($a in $PHAgents) { $weekInflow[$a] = 0; $weekClosure[$a] = 0 }

foreach ($row in $allCases) {
    $name = $row.PROCESSOR_NAME.Trim()
    if ($name -notin $PHAgents) { continue }
    # Inflow
    try { $created = [datetime]::Parse($row.CREATED_ON_DATE) } catch { continue }
    if ($created -ge $WeekStart -and $created -le $Today) { $weekInflow[$name]++ }
    # Closures
    $doneStr = if ($row.COMPLETED_ON -ne '') { $row.COMPLETED_ON } else { $row.CLOSED_ON }
    if (-not [string]::IsNullOrWhiteSpace($doneStr)) {
        try { $done = [datetime]::Parse($doneStr) } catch { continue }
        if ($done -ge $WeekStart -and $done -le $Today) { $weekClosure[$name]++ }
    }
}

# 2. AGING SCORE: active cases per agent (excluding Customer Action), deduct by age bucket
$agingScore  = @{}
$agingDetail = @{}
foreach ($a in $PHAgents) { $agingScore[$a] = 100; $agingDetail[$a] = @{b1=0;b2=0;b3=0} }

$activeCases = $allCases | Where-Object { $_.STATUS -notin @('Closed','Completed','Customer Action') }
foreach ($row in $activeCases) {
    $name = $row.PROCESSOR_NAME.Trim()
    if ($name -notin $PHAgents) { continue }
    try { $created = [datetime]::Parse($row.CREATED_ON_DATE) } catch { continue }
    $age = ($Today - $created).Days
    if ($age -ge 8 -and $age -le 15)  { $agingScore[$name] -= 1; $agingDetail[$name].b1++ }
    elseif ($age -ge 16 -and $age -le 30) { $agingScore[$name] -= 2; $agingDetail[$name].b2++ }
    elseif ($age -gt 30)              { $agingScore[$name] -= 4; $agingDetail[$name].b3++ }
}

# 3. CSAT BONUS: positive surveys (T2B) received this week per agent
$csatBonus = @{}
foreach ($a in $PHAgents) { $csatBonus[$a] = 0 }

$t2b = @('Very Satisfied','Somewhat Satisfied')
foreach ($row in $allSurveys) {
    $name = $row.PROCESSOR_NAME.Trim()
    if ($name -notin $PHAgents) { continue }
    if ($row.RATING -notin $t2b) { continue }
    $dateStr = if ($row.RECORDEDDATE -ne '') { $row.RECORDEDDATE } else { $row.COMPLETED_ON_TIMESTAMP }
    if ([string]::IsNullOrWhiteSpace($dateStr)) { continue }
    try { $rd = [datetime]::Parse($dateStr) } catch { continue }
    if ($rd -ge $WeekStart -and $rd -le $Today) { $csatBonus[$name]++ }
}

# -- RANK AND SCORE --
# Closure count ranking
$crRanked = $PHAgents | ForEach-Object {
    $i = $weekInflow[$_]; $c = $weekClosure[$_]
    $rate = if ($i -gt 0) { [math]::Round($c / $i * 100, 1) } else { if ($c -gt 0) { 999 } else { 0 } }
    [PSCustomObject]@{ Name=$_; Inflow=$i; Closures=$c; Rate=$rate }
} | Sort-Object Closures -Descending

$crPts = @{}
$rank = 1
foreach ($r in $crRanked) { $crPts[$r.Name] = (14 - $rank); $rank++ }

# Aging score ranking
$agRanked = $PHAgents | ForEach-Object {
    [PSCustomObject]@{ Name=$_; AgingScore=$agingScore[$_] }
} | Sort-Object AgingScore -Descending

$agPts = @{}
$rank = 1
foreach ($r in $agRanked) { $agPts[$r.Name] = (14 - $rank); $rank++ }

# -- BUILD RESULTS --
$results = $PHAgents | ForEach-Object {
    $cr  = $crPts[$_]
    $ag  = $agPts[$_]
    $cs  = $csatBonus[$_]
    $tot = $cr + $ag + $cs
    $rate = if ($weekInflow[$_] -gt 0) { "$([math]::Round($weekClosure[$_]/$weekInflow[$_]*100))%" } else { if ($weekClosure[$_] -gt 0) {'n/a'} else {'0%'} }
    [PSCustomObject]@{
        Name=$_; Inflow=$weekInflow[$_]; Closures=$weekClosure[$_]; Rate=$rate
        AgingScore=$agingScore[$_]
        B1=$agingDetail[$_].b1; B2=$agingDetail[$_].b2; B3=$agingDetail[$_].b3
        CRPts=$cr; AgPts=$ag; CSATBonus=$cs; Total=$tot
    }
} | Sort-Object Total -Descending

# Assign ranks (handle ties)
$ranked = @()
$pos = 1
for ($i = 0; $i -lt $results.Count; $i++) {
    if ($i -gt 0 -and $results[$i].Total -eq $results[$i-1].Total) {
        $ranked += $results[$i] | Select-Object *, @{N='Rank';E={"T$($pos-1)"}}
    } else {
        $pos = $i + 1
        $ranked += $results[$i] | Select-Object *, @{N='Rank';E={$pos}}
    }
}

# -- HTML --
$reportDate = $Today.ToString("MMMM d, yyyy")
$weekLabel  = ($WeekStart.ToString('MMM d') + ' - ' + $WeekEnd.ToString('MMM d, yyyy'))

$rowsHtml = ""
$pos = 0
foreach ($r in $ranked) {
    $pos++
    $medal = switch ($pos) { 1 {'&#129351;'} 2 {'&#129352;'} 3 {'&#129353;'} default {$pos} }
    $rateColor = '#9ca3af'
    $agColor = if ($r.AgingScore -ge 90) { '#166534' } elseif ($r.AgingScore -ge 70) { '#854d0e' } else { '#7f1d1d' }

    $b1s = if ($r.B1 -ne 1) { 's' } else { '' }
    $b2s = if ($r.B2 -ne 1) { 's' } else { '' }
    $b3s = if ($r.B3 -ne 1) { 's' } else { '' }
    $b1txt = if ($r.B1 -gt 0) { "<span style='color:#854d0e'>$($r.B1) case$b1s (8-15d)</span>" } else { '' }
    $b2txt = if ($r.B2 -gt 0) { "<span style='color:#c2410c'>$($r.B2) case$b2s (16-30d)</span>" } else { '' }
    $b3txt = if ($r.B3 -gt 0) { "<span style='color:#991b1b'>$($r.B3) case$b3s (30+d)</span>" } else { '' }
    $agingBreakdown = (@($b1txt,$b2txt,$b3txt) | Where-Object { $_ -ne '' }) -join ', '
    if (-not $agingBreakdown) { $agingBreakdown = '<span style="color:#166534">Clean</span>' }

    $csatTxt = if ($r.CSATBonus -gt 0) { "+$($r.CSATBonus)" } else { '--' }
    $totalColor = if ($pos -le 2) { '#92400e' } else { '#1e3a5f' }
    $rowBg = if ($pos -le 2) { 'background:#fffbeb;' } else { '' }

    $rowsHtml += @"
<tr style="$rowBg">
  <td style="text-align:center;font-size:18px">$medal</td>
  <td style="font-weight:600;color:#1e3a5f">$($r.Name)</td>
  <td style="text-align:center">$($r.Inflow)</td>
  <td style="text-align:center">$($r.Closures)</td>
  <td style="text-align:center;font-weight:700;color:$agColor">$($r.AgingScore)</td>
  <td style="font-size:11px;color:#6b7280">$agingBreakdown</td>
  <td style="text-align:center;font-weight:700;color:#059669">$csatTxt</td>
  <td style="text-align:center;font-weight:700;color:#1d4ed8">$($r.CRPts)</td>
  <td style="text-align:center;font-weight:700;color:#7c3aed">$($r.AgPts)</td>
  <td style="text-align:center;font-weight:700;color:#d97706">$($r.CSATBonus)</td>
  <td style="text-align:center;font-size:16px;font-weight:800;color:$totalColor">$($r.Total)</td>
</tr>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Backlog Busters – Week 1 Midweek Update</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a1a2e; }

    .header { background: linear-gradient(135deg, #1e3a5f, #2e5480); padding: 24px 32px; }
    .header h1 { font-size: 22px; font-weight: 800; color: #fff; margin-bottom: 4px; }
    .header .sub { font-size: 13px; color: #93c5fd; }
    .chips { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 12px; }
    .chip { background: rgba(255,255,255,0.12); border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #e0f0ff; }
    .chip.gold { background: rgba(240,171,0,0.2); color: #fde68a; border: 1px solid rgba(240,171,0,0.3); }

    .kpi-row { display: flex; gap: 16px; padding: 20px 32px; flex-wrap: wrap; }
    .kpi { background: #fff; border-radius: 12px; padding: 16px 22px; flex: 1; min-width: 140px; border-left: 4px solid #1e3a5f; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
    .kpi .label { font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
    .kpi .value { font-size: 28px; font-weight: 800; color: #1e3a5f; }
    .kpi .note  { font-size: 11px; color: #9ca3af; margin-top: 2px; }

    .section { padding: 0 32px 24px; }
    .section-title { font-size: 13px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 12px; padding-top: 8px; }

    table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.06); font-size: 13px; }
    thead tr { background: #1e3a5f; }
    thead th { padding: 10px 10px; color: #bfdbfe; font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.4px; border-right: 1px solid #2e5480; }
    thead th:last-child { border-right: none; }
    tbody tr { border-bottom: 1px solid #f3f4f6; transition: background 0.15s; }
    tbody tr:hover { background: #eff6ff !important; }
    tbody td { padding: 10px 10px; border-right: 1px solid #f3f4f6; }
    tbody td:last-child { border-right: none; }

    .th-group { background: #162d4a; font-size: 10px; color: #7dd3fc; text-align: center; padding: 4px; letter-spacing: 0.5px; }

    .legend { display: flex; gap: 24px; padding: 12px 32px 20px; font-size: 11px; color: #6b7280; flex-wrap: wrap; }
    .li { display: flex; align-items: center; gap: 6px; }
    .dot { width: 10px; height: 10px; border-radius: 50%; }

    .footer { padding: 16px 32px; font-size: 11px; color: #9ca3af; border-top: 1px solid #e5e7eb; }
  </style>
</head>
<body>

<div class="header">
  <h1>&#127942; Backlog Busters — Week 1 Midweek Update</h1>
  <div class="sub">PH Team Standings · As of $reportDate</div>
  <div class="chips">
    <div class="chip">&#128197; Week 1: $weekLabel</div>
    <div class="chip">&#127381; Results announced: Jun 30 (Tue)</div>
    <div class="chip gold">&#11088; 2 winners this week</div>
  </div>
</div>

<div class="kpi-row">
  <div class="kpi" style="border-color:#1d4ed8">
    <div class="label">Week Inflow</div>
    <div class="value" style="color:#1d4ed8">$( ($ranked | Measure-Object Inflow -Sum).Sum )</div>
    <div class="note">Jun 22–$( $Today.ToString('MMM d') )</div>
  </div>
  <div class="kpi" style="border-color:#059669">
    <div class="label">Week Closures</div>
    <div class="value" style="color:#059669">$( ($ranked | Measure-Object Closures -Sum).Sum )</div>
    <div class="note">Jun 22–$( $Today.ToString('MMM d') ) · PH team</div>
  </div>
  <div class="kpi" style="border-color:#d97706">
    <div class="label">CSAT Bonuses</div>
    <div class="value" style="color:#d97706">$( ($ranked | Measure-Object CSATBonus -Sum).Sum )</div>
    <div class="note">Positive surveys this week</div>
  </div>
  <div class="kpi" style="border-color:#7c3aed">
    <div class="label">Days Elapsed</div>
    <div class="value" style="color:#7c3aed">$( ($Today - $WeekStart).Days + 1 ) / 7</div>
    <div class="note">Week 1 progress</div>
  </div>
</div>

<div class="section">
  <div class="section-title">&#127942; Current Standings — PH Team (13 Agents)</div>
  <table>
    <thead>
      <tr>
        <th class="th-group" colspan="2"></th>
        <th class="th-group" colspan="2" style="border-left:1px solid #2e5480">This Week</th>
        <th class="th-group" colspan="2" style="border-left:1px solid #2e5480">Aging</th>
        <th class="th-group" colspan="1" style="border-left:1px solid #2e5480">CSAT</th>
        <th class="th-group" colspan="4" style="border-left:1px solid #2e5480;color:#F0AB00">Points Breakdown</th>
      </tr>
      <tr>
        <th style="width:40px">#</th>
        <th style="text-align:left;min-width:180px">Agent</th>
        <th>Inflow</th>
        <th>Closed</th>
        <th>Score</th>
        <th style="min-width:180px;text-align:left">Aging Detail</th>
        <th>Surveys</th>
        <th style="color:#93c5fd">CC Pts</th>
        <th style="color:#c4b5fd">Aging Pts</th>
        <th style="color:#fde68a">CSAT Pts</th>
        <th style="color:#F0AB00;font-size:13px">TOTAL</th>
      </tr>
    </thead>
    <tbody>
$rowsHtml
    </tbody>
  </table>
</div>

<div class="legend">
  <strong>Scoring:</strong>
  <div class="li"><div class="dot" style="background:#1d4ed8"></div> CC Pts &mdash; ranked by closure count (Rank 1 = most closures = 13 pts)</div>
  <div class="li"><div class="dot" style="background:#7c3aed"></div> Aging Pts &mdash; start 100, deduct per aged case (excl. Customer Action), then rank (Rank 1 = 13 pts)</div>
  <div class="li"><div class="dot" style="background:#d97706"></div> CSAT Bonus &mdash; +1 per positive survey received this week</div>
  <strong style="margin-left:8px">Aging deductions:</strong>
  <div class="li">8&ndash;15d = &minus;1</div>
  <div class="li">16&ndash;30d = &minus;2</div>
  <div class="li">30+d = &minus;4</div>
</div>

<div class="footer">
  Generated $reportDate &middot; Cases: $(Split-Path $CasesXlsx -Leaf) &middot; Surveys: $(Split-Path $SurveysXlsx -Leaf) &middot; Week 1 ends Jun 28 &middot; Final results announced Jun 30
</div>

</body>
</html>
"@

$outputFile = "$OutputDir\GTS_BacklogBusters_W1_Midweek.html"
$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! Saved to: $outputFile" -ForegroundColor Green
Start-Process $outputFile
