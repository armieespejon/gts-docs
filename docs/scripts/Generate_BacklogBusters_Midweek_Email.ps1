# ============================================================
# GTS Backlog Busters - Midweek Email Generator
# Scores PH team and opens a pre-filled .eml in Outlook
# ============================================================

$CasesXlsx   = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx" | Where-Object { $_.Name -notmatch 'with closed' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$SurveysXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Surveys_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$ReportPath  = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\GTS_BacklogBusters_W1_Midweek.html"

$To  = "your.recipient@sap.com"   # update before sending
$CC  = ""

# Week 1: Jun 22-28
$WeekStart = [datetime]'2026-06-22'
$WeekEnd   = [datetime]'2026-06-28'
$Today     = [datetime]::Today

Write-Host "Cases:   $CasesXlsx" -ForegroundColor Cyan
Write-Host "Surveys: $SurveysXlsx" -ForegroundColor Cyan

$PHAgents = @(
    'Nobela, Monica Frances','Quinto, Marviely','Chua, Sam','Santos, Ana Lizelle',
    'Mendoza, Edrian','Tulagan, Nicole','Valdez, Lemuel Von','Cruz, Ann',
    'Sinag, Rendel Edison','Evidente, Marie Cristine','Buenafe, Paul',
    'Cada, Saldy','Paulino, Justin'
)

# -- EXPORT --
$CasesCsv   = "$env:TEMP\gts_bb_email_cases.csv"
$SurveysCsv = "$env:TEMP\gts_bb_email_surveys.csv"
function Export-XlsxToCsv($xlsxPath, $csvPath) {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($xlsxPath)
    $wb.SaveAs($csvPath, 6)
    $wb.Close($false); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Write-Host "Exporting data..." -ForegroundColor Yellow
Export-XlsxToCsv $CasesXlsx $CasesCsv
Export-XlsxToCsv $SurveysXlsx $SurveysCsv

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','COMPLETED_ON','CLOSED_ON','IRT_HOURS')
$allCases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

$sHdr = @('DISPLAYID','CASETYPE','CASETYPEDESC','MONTH_DESC','EXT_COURSECODE','EXT_MISCINFO','SUBJECT','PRIORITY','SOURCE','SERVICETEAMNAME','ACCOUNTNAME','COMPLETED_ON_TIMESTAMP','RATING','COMMENT1','COMMENT2','RECORDEDDATE','PROCESSOR_ID','PROCESSOR_NAME','CAT1','CAT2','DOMAIN','FORWARDED_TO','COL23','COL24')
$allSurveys = Import-Csv $SurveysCsv -Header $sHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

# -- SCORING --
$weekInflow  = @{}; $weekClosure = @{}
foreach ($a in $PHAgents) { $weekInflow[$a] = 0; $weekClosure[$a] = 0 }

foreach ($row in $allCases) {
    $name = $row.PROCESSOR_NAME.Trim()
    if ($name -notin $PHAgents) { continue }
    try { $created = [datetime]::Parse($row.CREATED_ON_DATE) } catch { continue }
    if ($created -ge $WeekStart -and $created -le $Today) { $weekInflow[$name]++ }
    $doneStr = if ($row.COMPLETED_ON -ne '') { $row.COMPLETED_ON } else { $row.CLOSED_ON }
    if (-not [string]::IsNullOrWhiteSpace($doneStr)) {
        try { $done = [datetime]::Parse($doneStr) } catch { continue }
        if ($done -ge $WeekStart -and $done -le $Today) { $weekClosure[$name]++ }
    }
}

$agingScore = @{}; $agingDetail = @{}
foreach ($a in $PHAgents) { $agingScore[$a] = 100; $agingDetail[$a] = @{b1=0;b2=0;b3=0} }
$activeCases = $allCases | Where-Object { $_.STATUS -notin @('Closed','Completed','Customer Action') }
foreach ($row in $activeCases) {
    $name = $row.PROCESSOR_NAME.Trim()
    if ($name -notin $PHAgents) { continue }
    try { $created = [datetime]::Parse($row.CREATED_ON_DATE) } catch { continue }
    $age = ($Today - $created).Days
    if ($age -ge 8 -and $age -le 15)        { $agingScore[$name] -= 1; $agingDetail[$name].b1++ }
    elseif ($age -ge 16 -and $age -le 30)   { $agingScore[$name] -= 2; $agingDetail[$name].b2++ }
    elseif ($age -gt 30)                    { $agingScore[$name] -= 4; $agingDetail[$name].b3++ }
}

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

$crRanked = $PHAgents | ForEach-Object {
    [PSCustomObject]@{ Name=$_; Inflow=$weekInflow[$_]; Closures=$weekClosure[$_] }
} | Sort-Object Closures -Descending
$crPts = @{}; $rank = 1
foreach ($r in $crRanked) { $crPts[$r.Name] = (14 - $rank); $rank++ }

$agRanked = $PHAgents | ForEach-Object {
    [PSCustomObject]@{ Name=$_; AgingScore=$agingScore[$_] }
} | Sort-Object AgingScore -Descending
$agPts = @{}; $rank = 1
foreach ($r in $agRanked) { $agPts[$r.Name] = (14 - $rank); $rank++ }

$results = $PHAgents | ForEach-Object {
    [PSCustomObject]@{
        Name=$_; Inflow=$weekInflow[$_]; Closures=$weekClosure[$_]
        AgingScore=$agingScore[$_]
        B1=$agingDetail[$_].b1; B2=$agingDetail[$_].b2; B3=$agingDetail[$_].b3
        CRPts=$crPts[$_]; AgPts=$agPts[$_]; CSATBonus=$csatBonus[$_]
        Total=($crPts[$_] + $agPts[$_] + $csatBonus[$_])
    }
} | Sort-Object Total -Descending

$ranked = @(); $pos = 1
for ($i = 0; $i -lt $results.Count; $i++) {
    if ($i -gt 0 -and $results[$i].Total -eq $results[$i-1].Total) {
        $ranked += $results[$i] | Select-Object *, @{N='Rank';E={"T$($pos-1)"}}
    } else {
        $pos = $i + 1
        $ranked += $results[$i] | Select-Object *, @{N='Rank';E={$pos}}
    }
}

# -- EMAIL DATA --
$reportDate = $Today.ToString("MMMM d, yyyy")
$weekLabel  = ($WeekStart.ToString('MMM d') + ' - ' + $WeekEnd.ToString('MMM d, yyyy'))
$daysElapsed = ($Today - $WeekStart).Days + 1
$totalClosures = ($ranked | Measure-Object Closures -Sum).Sum
$totalCSAT     = ($ranked | Measure-Object CSATBonus -Sum).Sum
$leader = $ranked[0]
$medals = @('&#129351;','&#129352;','&#129353;')

# Top 3 callout line
$top3Line = ($ranked | Select-Object -First 3 | ForEach-Object {
    $i = [array]::IndexOf($ranked, $_)
    $fn = ($_.Name -split ',')[1].Trim().Split(' ')[0]
    "$($medals[$i]) <strong>$fn</strong> ($($_.Total) pts)"
}) -join ' &nbsp;&nbsp; '

# Standings table rows
$tableRows = ""
$pos = 0
foreach ($r in $ranked) {
    $pos++
    $medal = switch ($pos) { 1 {'&#129351;'} 2 {'&#129352;'} 3 {'&#129353;'} default {$pos} }
    $rowBg  = if ($pos -le 2) { "background:#fffbeb;" } else { "" }
    $totCol = if ($pos -le 2) { "#92400e" } else { "#1e3a5f" }
    $agCol  = if ($r.AgingScore -ge 90) { "#166534" } elseif ($r.AgingScore -ge 70) { "#854d0e" } else { "#7f1d1d" }
    $b1txt  = if ($r.B1 -gt 0) { "<span style='color:#854d0e'>$($r.B1) (8-15d)</span>" } else { "" }
    $b2txt  = if ($r.B2 -gt 0) { "<span style='color:#c2410c'>$($r.B2) (16-30d)</span>" } else { "" }
    $b3txt  = if ($r.B3 -gt 0) { "<span style='color:#991b1b'>$($r.B3) (30+d)</span>" } else { "" }
    $aging  = (@($b1txt,$b2txt,$b3txt) | Where-Object { $_ -ne "" }) -join ", "
    if (-not $aging) { $aging = "<span style='color:#166534'>Clean</span>" }
    $csat   = if ($r.CSATBonus -gt 0) { "+$($r.CSATBonus)" } else { "--" }

    $tableRows += "
<tr style='${rowBg}border-bottom:1px solid #f3f4f6'>
  <td style='padding:7px 10px;text-align:center;font-size:16px'>$medal</td>
  <td style='padding:7px 10px;font-weight:600;color:#1e3a5f'>$($r.Name)</td>
  <td style='padding:7px 10px;text-align:center'>$($r.Inflow)</td>
  <td style='padding:7px 10px;text-align:center'>$($r.Closures)</td>
  <td style='padding:7px 10px;text-align:center;font-weight:700;color:$agCol'>$($r.AgingScore)</td>
  <td style='padding:7px 10px;font-size:11px;color:#6b7280'>$aging</td>
  <td style='padding:7px 10px;text-align:center;color:#059669;font-weight:700'>$csat</td>
  <td style='padding:7px 10px;text-align:center;color:#1d4ed8;font-weight:700'>$($r.CRPts)</td>
  <td style='padding:7px 10px;text-align:center;color:#7c3aed;font-weight:700'>$($r.AgPts)</td>
  <td style='padding:7px 10px;text-align:center;color:#d97706;font-weight:700'>$($r.CSATBonus)</td>
  <td style='padding:7px 10px;text-align:center;font-size:15px;font-weight:800;color:$totCol'>$($r.Total)</td>
</tr>"
}

$subject = "GTS | Backlog Busters W1 Midweek Update | $reportDate"

$body = @"
<html><body style="font-family:'Segoe UI',Arial,sans-serif;font-size:13px;color:#1a1a2e;max-width:780px">

<div style="background:linear-gradient(135deg,#1e3a5f,#2e5480);padding:16px 20px;border-radius:6px 6px 0 0">
  <div style="font-size:18px;font-weight:800;color:#fff">&#127942; Backlog Busters &mdash; Week 1 Midweek Update</div>
  <div style="font-size:12px;color:#93c5fd;margin-top:4px">PH Team Standings &middot; $reportDate &middot; Week 1: $weekLabel</div>
</div>

<div style="background:#f8fafc;padding:14px 20px;border:1px solid #e5e7eb;border-top:none">
  <p style="margin:0 0 6px">Hi team,</p>
  <p style="margin:0 0 6px">Here is the midweek Backlog Busters standings as of <strong>$reportDate</strong> (Day $daysElapsed of 7). Final results to be announced on <strong>June 30th</strong>.</p>
  <p style="margin:0 0 14px">Thank you for your hard work this week! &#128170; Let&rsquo;s take on that backlog &mdash; there&rsquo;s still time to catch up. Keep closing those cases, update your statuses, and watch out for aging. Every case counts! &#127942;</p>

  <!-- TOP 3 CALLOUT -->
  <div style="background:#F0AB00;border-radius:6px;padding:12px 16px;margin-bottom:16px">
    <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;color:#78350f;margin-bottom:7px">&#127942; Current Top 3</div>
    <div style="font-size:14px;font-weight:600;color:#1a1a2e">$top3Line</div>
  </div>

  <!-- KPI STRIP -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px">
    <tr>
      <td style="width:24%;padding:10px;background:#eff6ff;border-radius:6px;text-align:center;border:1px solid #bfdbfe">
        <div style="font-size:20px;font-weight:700;color:#1d4ed8">$daysElapsed / 7</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">Days Elapsed</div>
      </td>
      <td style="width:4%"></td>
      <td style="width:24%;padding:10px;background:#f0fdf4;border-radius:6px;text-align:center;border:1px solid #bbf7d0">
        <div style="font-size:20px;font-weight:700;color:#166534">$totalClosures</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">Week Closures (PH)</div>
      </td>
      <td style="width:4%"></td>
      <td style="width:24%;padding:10px;background:#fefce8;border-radius:6px;text-align:center;border:1px solid #fde68a">
        <div style="font-size:20px;font-weight:700;color:#92400e">$totalCSAT</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">CSAT Bonuses</div>
      </td>
      <td style="width:4%"></td>
      <td style="width:20%;padding:10px;background:#faf5ff;border-radius:6px;text-align:center;border:1px solid #e9d5ff">
        <div style="font-size:20px;font-weight:700;color:#7c3aed">$($leader.Total)</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">Leader Score</div>
      </td>
    </tr>
  </table>

  <!-- STANDINGS TABLE -->
  <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;font-size:12px">
    <thead>
      <tr style="background:#1e3a5f">
        <th colspan="2" style="padding:6px 10px;text-align:left;color:#fff;font-size:10px;letter-spacing:0.4px;border-right:1px solid #2e5480"></th>
        <th colspan="2" style="padding:6px 10px;text-align:center;color:#93c5fd;font-size:10px;letter-spacing:0.4px;border-right:1px solid #2e5480">THIS WEEK</th>
        <th colspan="2" style="padding:6px 10px;text-align:center;color:#a7f3d0;font-size:10px;letter-spacing:0.4px;border-right:1px solid #2e5480">AGING</th>
        <th style="padding:6px 10px;text-align:center;color:#fde68a;font-size:10px;letter-spacing:0.4px;border-right:1px solid #2e5480">CSAT</th>
        <th colspan="3" style="padding:6px 10px;text-align:center;color:#c4b5fd;font-size:10px;letter-spacing:0.4px;border-right:1px solid #2e5480">POINTS</th>
        <th style="padding:6px 10px;text-align:center;color:#F0AB00;font-size:11px;letter-spacing:0.4px">TOTAL</th>
      </tr>
      <tr style="background:#162d4a">
        <th style="padding:7px 10px;text-align:center;color:#bfdbfe;font-weight:600">#</th>
        <th style="padding:7px 10px;text-align:left;color:#fff;font-weight:600;min-width:160px">Agent</th>
        <th style="padding:7px 10px;text-align:center;color:#93c5fd;font-weight:600">Inflow</th>
        <th style="padding:7px 10px;text-align:center;color:#6ee7b7;font-weight:600">Closed</th>
        <th style="padding:7px 10px;text-align:center;color:#a7f3d0;font-weight:600">Score</th>
        <th style="padding:7px 10px;text-align:left;color:#a7f3d0;font-weight:600;min-width:140px">Aged Cases</th>
        <th style="padding:7px 10px;text-align:center;color:#fde68a;font-weight:600">Bonus</th>
        <th style="padding:7px 10px;text-align:center;color:#93c5fd;font-weight:600">CR</th>
        <th style="padding:7px 10px;text-align:center;color:#c4b5fd;font-weight:600">Ag</th>
        <th style="padding:7px 10px;text-align:center;color:#fde68a;font-weight:600">CS</th>
        <th style="padding:7px 10px;text-align:center;color:#F0AB00;font-weight:700">&#931;</th>
      </tr>
    </thead>
    <tbody>$tableRows
    </tbody>
  </table>

  <p style="margin:14px 0 0;font-size:11px;color:#6b7280">Open the attached HTML for the full interactive report.</p>
</div>

<div style="background:#f1f5f9;padding:10px 20px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;font-size:11px;color:#9ca3af">
  GTS Backlog Busters Sprint &middot; Jun 22 &ndash; Jul 19, 2026 &middot; Auto-generated &middot; $reportDate
</div>

</body></html>
"@

# -- GENERATE EML --
Write-Host "Generating email file..." -ForegroundColor Yellow
$emlPath = "$env:TEMP\GTS_BB_Midweek_$($Today.ToString('MMMdd')).eml"
$boundary = "----=_GTSBB_$(Get-Random)"
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
$bodyB64   = [Convert]::ToBase64String($bodyBytes)

$attachB64 = ""; $attachName = ""
if (Test-Path $ReportPath) {
    $attachB64  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ReportPath))
    $attachName = Split-Path $ReportPath -Leaf
}

$ccField = if ($CC) { "CC: $CC`r`n" } else { "" }
$eml  = "From: GTS ESM Operations`r`n"
$eml += "To: $To`r`n"
$eml += $ccField
$eml += "Subject: $subject`r`n"
$eml += "MIME-Version: 1.0`r`n"
$eml += "Content-Type: multipart/mixed; boundary=`"$boundary`"`r`n`r`n"
$eml += "--$boundary`r`nContent-Type: text/html; charset=UTF-8`r`nContent-Transfer-Encoding: base64`r`n`r`n$bodyB64`r`n"
if ($attachB64) {
    $eml += "--$boundary`r`nContent-Type: text/html; name=`"$attachName`"`r`nContent-Transfer-Encoding: base64`r`nContent-Disposition: attachment; filename=`"$attachName`"`r`n`r`n$attachB64`r`n"
}
$eml += "--$boundary--`r`n"

[System.IO.File]::WriteAllText($emlPath, $eml, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "Done! Email file saved. Opening now..." -ForegroundColor Green
Write-Host "  $emlPath" -ForegroundColor Yellow
Start-Process $emlPath
