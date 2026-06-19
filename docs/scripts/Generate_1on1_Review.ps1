# ============================================================
# GTS Team Manager 1:1 Review -- Auto-Generator
# Reads cases CSV, computes all agent stats, builds HTML report
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_1on1_Review.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$CasesXlsx  = "C:\Users\I535893\Downloads\GTS ESM Cases_June17.xlsx"
$OutputDir  = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"
$ReportDate = "June 17, 2026"
$Today      = [datetime]'2026-06-17'
$TargetClosureRate = 90
# -- END CONFIG -----------------------------------------------

$CasesCsv = "$env:TEMP\gts_cases_1on1.csv"

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

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE','EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13','L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY','CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing agent stats..." -ForegroundColor Yellow

# Known agent names (filter to main team only)
$teamNames = @(
    'Nobela, Monica Frances','Quinto, Marviely','Espejon, Armie','Chua, Sam',
    'Santos, Ana Lizelle','Tulagan, Nicole','Mendoza, Edrian','Cruz, Ann',
    'Valdez, Lemuel Von','Sinag, Rendel Edison','Mritunjay Kumar Singh','Shubham Thakur',
    'Evidente, Marie Cristine','Buenafe, Paul','Anjali Tripathi','Swarna Gupta',
    'Kriti Bhalla','Cada, Saldy','Paulino, Justin','Narendran Raja'
)

function Resolve-Name($row) {
    if ($row.PROCESSOR_ID -eq 'SAP_GERMANY_USER') {
        $misc = $row.EXT_MISCINFO.Trim().ToLower()
        if ($misc -match 'mjay')   { return 'Mritunjay Kumar Singh' }
        if ($misc -match 'shub')   { return 'Shubham Thakur' }
        if ($misc -match 'ani')    { return 'Anjali Tripathi' }
        if ($misc -match 'swarna') { return 'Swarna Gupta' }
        if ($misc -match 'kriti')  { return 'Kriti Bhalla' }
        if ($misc -match 'naren')  { return 'Narendran Raja' }
        return $null
    }
    $n = $row.PROCESSOR_NAME.Trim()
    if ([string]::IsNullOrWhiteSpace($n) -or $n -eq 'None') { return $null }
    return $n
}

$agentData = @{}
foreach ($name in $teamNames) {
    $agentData[$name] = @{ YTDIn=0; YTDClosed=0; Active=@() }
}

foreach ($row in $cases) {
    $name = Resolve-Name $row
    if ($null -eq $name -or -not $agentData.ContainsKey($name)) { continue }
    $agentData[$name].YTDIn++
    if ($row.STATUS -in @('Closed','Completed')) {
        $agentData[$name].YTDClosed++
    } else {
        $agentData[$name].Active += $row
    }
}

# Compute derived stats per agent
$agentStats = @()
foreach ($name in $teamNames) {
    $a = $agentData[$name]
    $rate = if ($a.YTDIn -gt 0) { [math]::Round($a.YTDClosed / $a.YTDIn * 100) } else { 0 }

    $d07=0; $d815=0; $d1630=0; $d30p=0; $ages=@()
    foreach ($row in $a.Active) {
        try {
            $age = ($Today - [datetime]::Parse($row.CREATED_ON_DATE)).Days
            $ages += $age
            if     ($age -le 7)  { $d07++ }
            elseif ($age -le 15) { $d815++ }
            elseif ($age -le 30) { $d1630++ }
            else                 { $d30p++ }
        } catch {}
    }
    $avgAge = if ($ages.Count -gt 0) { [math]::Round(($ages | Measure-Object -Average).Average, 1) } else { 0 }
    $maxAge = if ($ages.Count -gt 0) { ($ages | Measure-Object -Maximum).Maximum } else { 0 }

    $agentStats += [PSCustomObject]@{
        Name=$name; YTDIn=$a.YTDIn; YTDClosed=$a.YTDClosed; Rate=$rate
        Active=$a.Active.Count; D07=$d07; D815=$d815; D1630=$d1630; D30p=$d30p
        AvgAge=$avgAge; MaxAge=$maxAge
    }
}

# Team totals
$tIn     = ($agentStats | Measure-Object YTDIn -Sum).Sum
$tClosed = ($agentStats | Measure-Object YTDClosed -Sum).Sum
$tActive = ($agentStats | Measure-Object Active -Sum).Sum
$tRate   = [math]::Round($tClosed / $tIn * 100)
$t07     = ($agentStats | Measure-Object D07 -Sum).Sum
$t815    = ($agentStats | Measure-Object D815 -Sum).Sum
$t1630   = ($agentStats | Measure-Object D1630 -Sum).Sum
$t30p    = ($agentStats | Measure-Object D30p -Sum).Sum

# Categorize agents
$needsAttn  = $agentStats | Where-Object { $_.D30p -gt 0 -or $_.Rate -lt $TargetClosureRate } | Sort-Object { $_.D30p * 100 + $_.D1630 } -Descending
$performing = $agentStats | Where-Object { $_.D30p -eq 0 -and $_.Rate -ge 95 } | Sort-Object Rate -Descending

Write-Host "  Team: $tIn in, $tClosed closed ($tRate%), $tActive active" -ForegroundColor Gray

# ── Build HTML functions ──────────────────────────────────────

function Get-RateColor($rate) {
    if ($rate -ge 96) { return '#047857' }
    if ($rate -ge 93) { return '#1d4ed8' }
    if ($rate -ge 90) { return '#b45309' }
    return '#b91c1c'
}

function Get-AgingColor($val, $type) {
    if ($val -eq 0) { return '#94a3b8' }
    if ($type -eq '30p') { return '#b91c1c' }
    if ($type -eq '1630') { return '#d97706' }
    return '#1e293b'
}

function Build-AgentRow($s) {
    $rateColor = Get-RateColor $s.Rate
    $c30  = if ($s.D30p  -gt 0) { '#b91c1c' } else { '#94a3b8' }
    $c16  = if ($s.D1630 -gt 0) { '#d97706' } else { '#94a3b8' }
    return @"
            <tr>
              <td class="name-col">$($s.Name)</td>
              <td>$($s.YTDIn.ToString('N0'))</td>
              <td>$($s.YTDClosed.ToString('N0'))</td>
              <td style="color:$rateColor;font-weight:700;">$($s.Rate)%</td>
              <td style="font-weight:600;">$($s.Active)</td>
              <td>$($s.D07)</td>
              <td>$($s.D815)</td>
              <td style="color:$c16;font-weight:$(if($s.D1630 -gt 0){'700'}else{'400'});">$($s.D1630)</td>
              <td style="color:$c30;font-weight:$(if($s.D30p -gt 0){'700'}else{'400'});">$($s.D30p)</td>
              <td>$($s.AvgAge)</td>
              <td>$($s.MaxAge)</td>
            </tr>
"@
}

function Build-NeedsAttnRow($s) {
    $c30  = if ($s.D30p  -gt 0) { 'style="color:#b91c1c;font-weight:700;"' } else { '' }
    $c16  = if ($s.D1630 -gt 0) { 'style="color:#d97706;font-weight:700;"' } else { '' }
    $cr   = if ($s.Rate -lt $TargetClosureRate) { 'style="color:#b91c1c;font-weight:700;"' } else { 'style="color:#1d4ed8;font-weight:700;"' }
    return @"
            <tr>
              <td class="name-col" style="font-weight:700;">$($s.Name)</td>
              <td>$($s.Active)</td>
              <td $c30>$($s.D30p)</td>
              <td $c16>$($s.D1630)</td>
              <td $cr>$($s.Rate)%</td>
            </tr>
"@
}

function Build-PerformingRow($s) {
    return @"
            <tr>
              <td class="name-col" style="font-weight:700;color:#047857;">$($s.Name)</td>
              <td>$($s.Active)</td>
              <td>$($s.AvgAge)</td>
              <td style="color:#047857;font-weight:700;">$($s.Rate)%</td>
            </tr>
"@
}

function Build-AgentPoints($s) {
    $pts = ""
    if ($s.D30p -gt 0) {
        $pts += "<li><strong style='color:#b91c1c;'>$($s.D30p) case(s) older than 30 days</strong> -- must be resolved or escalated.</li>"
    }
    if ($s.D1630 -gt 0) {
        $pts += "<li><strong style='color:#d97706;'>$($s.D1630) case(s) in the 16-30 day bracket</strong> -- action required this week.</li>"
    }
    if ($s.Rate -lt $TargetClosureRate) {
        $pts += "<li><strong style='color:#b91c1c;'>Closure rate is $($s.Rate)%</strong> -- below the $TargetClosureRate% team target. Review blockers.</li>"
    }
    if ($pts -eq "") {
        $pts = "<li style='color:#047857;'>No critical items -- keep maintaining current performance.</li>"
    }
    return $pts
}

function Get-AgingBadge($val, $type) {
    $col = switch ($type) {
        '07'   { '#1e3a5f' }
        '815'  { '#1d4ed8' }
        '1630' { if ($val -gt 0) { '#d97706' } else { '#94a3b8' } }
        '30p'  { if ($val -gt 0) { '#b91c1c' } else { '#94a3b8' } }
    }
    $bg = switch ($type) {
        '07'   { '#eff6ff' }
        '815'  { '#dbeafe' }
        '1630' { if ($val -gt 0) { '#fef3c7' } else { '#f8fafc' } }
        '30p'  { if ($val -gt 0) { '#fee2e2' } else { '#f8fafc' } }
    }
    return "<div class='age-box' style='background:$bg;color:$col;'><div class='age-val'>$val</div><div class='age-lbl'>$(switch($type){'07'{'0-7 Days'}'815'{'8-15 Days'}'1630'{'16-30 Days'}'30p'{'>30 Days'}})</div></div>"
}

# Build HTML sections
$needsAttnRows = ($needsAttn | ForEach-Object { Build-NeedsAttnRow $_ }) -join ""
$performingRows = ($performing | ForEach-Object { Build-PerformingRow $_ }) -join ""
$allAgentRows = ($agentStats | ForEach-Object { Build-AgentRow $_ }) -join ""

# Build individual 1:1 sheets
$agentSheets = ""
foreach ($s in $agentStats) {
    $rateColor = Get-RateColor $s.Rate
    $points    = Build-AgentPoints $s
    $b07  = Get-AgingBadge $s.D07   '07'
    $b815 = Get-AgingBadge $s.D815  '815'
    $b16  = Get-AgingBadge $s.D1630 '1630'
    $b30  = Get-AgingBadge $s.D30p  '30p'

    $agentSheets += @"
  <div class="agent-sheet">
    <div class="agent-header">
      <div>
        <h2 class="agent-name">$($s.Name)</h2>
        <div class="agent-sub">1:1 Case Review | $ReportDate</div>
      </div>
    </div>

    <div class="kpi-strip">
      <div class="kpi-item"><div class="kpi-lbl">YTD Inflow</div><div class="kpi-val blue">$($s.YTDIn.ToString('N0'))</div></div>
      <div class="kpi-item"><div class="kpi-lbl">YTD Closed</div><div class="kpi-val blue">$($s.YTDClosed.ToString('N0'))</div></div>
      <div class="kpi-item"><div class="kpi-lbl">YTD Closure Rate</div><div class="kpi-val" style="color:$rateColor;">$($s.Rate)%</div></div>
      <div class="kpi-item"><div class="kpi-lbl">Active Cases</div><div class="kpi-val blue">$($s.Active)</div></div>
      <div class="kpi-item"><div class="kpi-lbl">Avg Age</div><div class="kpi-val $(if($s.AvgAge -gt 20){'red'}elseif($s.AvgAge -gt 12){'amber'}else{'blue'})">$($s.AvgAge)</div></div>
    </div>

    <div class="section-label">Active Cases by Age</div>
    <div class="age-row">
      $b07 $b815 $b16 $b30
    </div>

    <div class="section-label">Points to Discuss</div>
    <ul class="points-list">$points</ul>

    <div class="section-label">1:1 Notes</div>
    <div class="notes-block"><div class="notes-label">Agent Update / Comments:</div><div class="notes-lines"></div></div>
    <div class="notes-block"><div class="notes-label">Actions Agreed:</div><div class="notes-lines"></div></div>
    <div class="notes-block"><div class="notes-label">Follow-up Required:</div><div class="notes-lines"></div></div>
    <div class="notes-block"><div class="notes-label">Coaching Log Entry:</div><div class="notes-lines"></div></div>

    <div class="sheet-footer">GTS Management &nbsp;|&nbsp; 1:1 Review &nbsp;|&nbsp; $($s.Name) &nbsp;|&nbsp; $ReportDate</div>
  </div>
"@
}

# ── Write HTML ────────────────────────────────────────────────
$outputFile = "$OutputDir\GTS_Manager_1on1_Review_Jun17.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>GTS Team -- Manager 1:1 Review | $ReportDate</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1e293b; }

    /* PAGE HEADER */
    .page-header { background: #1e3a5f; padding: 20px 32px 16px; }
    .page-header h1 { font-size: 22px; font-weight: 700; color: #fff; margin-bottom: 4px; }
    .page-header p  { font-size: 12px; color: #93c5fd; }

    /* MANAGER SUMMARY */
    .summary { background: #fff; margin: 20px 32px; border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
    .summary h2 { font-size: 14px; font-weight: 700; color: #1e3a5f; margin-bottom: 16px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }

    /* SNAPSHOT TILES */
    .snapshot { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 24px; }
    .snap-tile { background: #f8fafc; border-radius: 8px; padding: 14px 16px; border-top: 3px solid var(--tc); }
    .snap-tile.blue { --tc: #1d4ed8; } .snap-tile.green { --tc: #047857; } .snap-tile.amber { --tc: #d97706; } .snap-tile.red { --tc: #b91c1c; }
    .snap-lbl { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #64748b; margin-bottom: 6px; }
    .snap-val { font-size: 30px; font-weight: 700; color: var(--tc); }

    .aging-strip { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 24px; }
    .aging-tile { background: #f8fafc; border-radius: 8px; padding: 12px 16px; text-align: center; }
    .aging-num { font-size: 26px; font-weight: 700; }
    .aging-lbl { font-size: 11px; color: #64748b; margin-top: 2px; }
    .aging-tile.a07  .aging-num { color: #1e3a5f; }
    .aging-tile.a815 .aging-num { color: #1d4ed8; }
    .aging-tile.a16  .aging-num { color: #d97706; }
    .aging-tile.a30  .aging-num { color: #b91c1c; font-size: 32px; }

    /* WATCH LIST TABLES */
    .watch-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 24px; }
    .watch-panel h3 { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 10px; }
    .watch-panel h3.red { color: #b91c1c; }
    .watch-panel h3.green { color: #047857; }
    .watch-sub { font-size: 11px; color: #64748b; margin-bottom: 8px; }

    /* TABLES */
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    thead tr { background: #1e3a5f; color: #fff; }
    thead th { padding: 8px 8px; text-align: center; font-weight: 600; font-size: 11px; white-space: nowrap; }
    thead th.name-col { text-align: left; padding-left: 12px; }
    tbody tr { border-bottom: 1px solid #f1f5f9; }
    tbody tr:hover { background: #f8fafc; }
    tbody td { padding: 7px 8px; text-align: center; color: #374151; }
    tbody td.name-col { text-align: left; padding-left: 12px; font-weight: 500; }

    /* FULL TEAM TABLE */
    .full-table-wrap { overflow-x: auto; }

    /* AGENT SHEETS */
    .agent-sheet { background: #fff; margin: 20px 32px; border-radius: 10px; padding: 28px 32px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); page-break-before: always; }
    .agent-header { margin-bottom: 20px; padding-bottom: 14px; border-bottom: 2px solid #1e3a5f; }
    .agent-name { font-size: 22px; font-weight: 700; color: #1d4ed8; }
    .agent-sub  { font-size: 12px; color: #64748b; margin-top: 3px; }

    .kpi-strip { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 20px; }
    .kpi-item { background: #f8fafc; border-radius: 8px; padding: 12px 14px; text-align: center; border-top: 3px solid #1d4ed8; }
    .kpi-lbl { font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #64748b; margin-bottom: 6px; }
    .kpi-val { font-size: 24px; font-weight: 700; }
    .kpi-val.blue  { color: #1d4ed8; }
    .kpi-val.red   { color: #b91c1c; }
    .kpi-val.amber { color: #d97706; }

    .section-label { font-size: 11px; font-weight: 700; color: #1e3a5f; text-transform: uppercase; letter-spacing: 0.8px; margin: 16px 0 10px; }

    .age-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px; }
    .age-box { border-radius: 8px; padding: 12px 14px; text-align: center; }
    .age-val { font-size: 28px; font-weight: 700; line-height: 1; }
    .age-lbl { font-size: 10px; font-weight: 600; margin-top: 4px; opacity: 0.8; }

    .points-list { margin-left: 18px; margin-bottom: 16px; }
    .points-list li { font-size: 12px; color: #374151; line-height: 1.8; }

    .notes-block { margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px solid #e2e8f0; }
    .notes-block:last-of-type { border-bottom: none; }
    .notes-label { font-size: 12px; font-weight: 600; color: #1e293b; margin-bottom: 32px; }
    .notes-lines { height: 60px; border-bottom: 1px solid #e2e8f0; margin-top: 8px; }

    .sheet-footer { text-align: center; font-size: 10px; color: #94a3b8; margin-top: 24px; padding-top: 12px; border-top: 1px solid #e2e8f0; }

    @media print {
      .agent-sheet { page-break-before: always; margin: 0; border-radius: 0; box-shadow: none; }
      .summary { page-break-after: always; margin: 0; border-radius: 0; box-shadow: none; }
    }
  </style>
</head>
<body>

<div class="page-header">
  <h1>GTS Team -- Manager Summary &amp; 1:1 Review</h1>
  <p>As of $ReportDate &nbsp;|&nbsp; GTS Management &nbsp;|&nbsp; Confidential</p>
</div>

<!-- MANAGER SUMMARY -->
<div class="summary">
  <h2>Team Performance Snapshot</h2>

  <div class="snapshot">
    <div class="snap-tile blue"><div class="snap-lbl">YTD Inflow</div><div class="snap-val">$($tIn.ToString('N0'))</div></div>
    <div class="snap-tile green"><div class="snap-lbl">YTD Closures</div><div class="snap-val">$($tClosed.ToString('N0'))</div></div>
    <div class="snap-tile $(if($tRate -ge 93){'green'}elseif($tRate -ge 90){'amber'}else{'red'})"><div class="snap-lbl">YTD Closure Rate</div><div class="snap-val">$tRate%</div></div>
    <div class="snap-tile amber"><div class="snap-lbl">Active Cases</div><div class="snap-val">$($tActive.ToString('N0'))</div></div>
  </div>

  <div class="aging-strip">
    <div class="aging-tile a07"><div class="aging-num">$t07</div><div class="aging-lbl">0-7 Days</div></div>
    <div class="aging-tile a815"><div class="aging-num">$t815</div><div class="aging-lbl">8-15 Days</div></div>
    <div class="aging-tile a16"><div class="aging-num">$t1630</div><div class="aging-lbl">16-30 Days</div></div>
    <div class="aging-tile a30"><div class="aging-num">$t30p</div><div class="aging-lbl">&gt;30 Days</div></div>
  </div>

  <h2>Highlights &amp; Watch List</h2>

  <div class="watch-grid">
    <div class="watch-panel">
      <h3 class="red">Needs Attention</h3>
      <div class="watch-sub">Agents with &gt;30-day cases or closure rate below $TargetClosureRate%</div>
      <table>
        <thead><tr><th class="name-col">Agent</th><th>Active</th><th>&gt;30d Cases</th><th>16-30d Cases</th><th>Closure Rate</th></tr></thead>
        <tbody>$needsAttnRows</tbody>
      </table>
    </div>
    <div class="watch-panel">
      <h3 class="green">Performing Well</h3>
      <div class="watch-sub">Agents with 95%+ closure rate and no &gt;30-day cases</div>
      <table>
        <thead><tr><th class="name-col">Agent</th><th>Active</th><th>Avg Age</th><th>Closure Rate</th></tr></thead>
        <tbody>$performingRows</tbody>
      </table>
    </div>
  </div>

  <h2>Full Team Overview</h2>
  <div class="full-table-wrap">
    <table>
      <thead>
        <tr>
          <th class="name-col">Agent</th>
          <th>YTD In</th><th>YTD Closed</th><th>YTD Rate</th>
          <th>Active</th><th>0-7d</th><th>8-15d</th><th>16-30d</th><th>&gt;30d</th>
          <th>Avg Age</th><th>Max Age</th>
        </tr>
      </thead>
      <tbody>$allAgentRows</tbody>
      <tfoot>
        <tr style="background:#1e3a5f;color:#fff;font-weight:700;">
          <td class="name-col" style="padding-left:12px;">TEAM TOTAL</td>
          <td>$($tIn.ToString('N0'))</td><td>$($tClosed.ToString('N0'))</td><td>$tRate%</td>
          <td>$($tActive.ToString('N0'))</td><td>$t07</td><td>$t815</td><td>$t1630</td><td>$t30p</td>
          <td>-</td><td>-</td>
        </tr>
      </tfoot>
    </table>
  </div>
</div>

<!-- INDIVIDUAL 1:1 SHEETS -->
$agentSheets

</body>
</html>
"@

$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
