# ============================================================
# GTS ESM Executive Summary - Auto-Generator
#
# Usage:
#   1. Drop updated xlsx files in Downloads (or update paths below)
#   2. Update CONFIG section (dates, targets, prior year figures)
#   3. Run: powershell -ExecutionPolicy Bypass -File "Generate_Exec_Summary.ps1"
# ============================================================

# ?? CONFIG ???????????????????????????????????????????????????
$SurveysXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Surveys_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$CasesXlsx   = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx"   | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$OutputDir   = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"

# Derive report date from filename (e.g. "GTS ESM Cases_June18.xlsx" -> June 18, 2026)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($CasesXlsx)
$datePart  = $baseName -replace '^GTS ESM Cases_', ''
$parsedDate = $null
foreach ($fmt in @('MMMMd','MMMMdd','MMMd','MMMdd')) {
    try {
        $parsedDate = [datetime]::ParseExact($datePart, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
        break
    } catch {}
}
if (-not $parsedDate) { $parsedDate = [datetime]::Today }
$parsedDate = $parsedDate.AddYears([datetime]::Today.Year - $parsedDate.Year)

$ReportDate        = $parsedDate.ToString("MMMM d, yyyy")
$ReportDateShort   = $parsedDate.ToString("MMM d")
$CurrentMonth      = $parsedDate.ToString("MMMM")
$CurrentMonthShort = $parsedDate.ToString("MMM")
$YtdLabel          = "Jan - $ReportDateShort, 2026"

# Prior year benchmarks
$PY_Volume   = 9135
$PY_SBA      = 819
$PY_CSAT_PCT = 84.18
$PY_CSAT_N   = 1846

# CSAT target
$CSAT_TARGET = 80

# Tip / goals for footer (optional - leave blank to omit)
# ?? END CONFIG ???????????????????????????????????????????????

$CasesCsv   = "$env:TEMP\gts_cases_exec.csv"
$SurveysCsv = "$env:TEMP\gts_surveys_exec.csv"

# ?? Export Excel ? CSV ???????????????????????????????????????
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
Export-XlsxToCsv $CasesXlsx   $CasesCsv
Export-XlsxToCsv $SurveysXlsx $SurveysCsv

# ?? Load cases ???????????????????????????????????????????????
$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }
$today = [datetime]::Today

Write-Host "Computing case metrics..." -ForegroundColor Yellow

# KPI: volumes
$ytdTotal  = $cases.Count
$active    = ($cases | Where-Object { $_.STATUS -notin @('Closed','Completed') }).Count
$closed    = $ytdTotal - $active

# Load previous active count from state file, then save today's
$stateFile = "$OutputDir\exec_state.json"
$prevActive = $null
$prevDate   = $null
if (Test-Path $stateFile) {
    $state = Get-Content $stateFile | ConvertFrom-Json
    $prevActive = $state.ActiveCases
    $prevDate   = $state.ReportDate
}
@{ ActiveCases = $active; ReportDate = $ReportDate } | ConvertTo-Json | Out-File $stateFile -Encoding UTF8

# Build active delta note
if ($prevActive -ne $null -and $prevDate -ne $null) {
    $activeDiff = $active - $prevActive
    $activeDiffFmt = if ($activeDiff -gt 0) { "+$activeDiff" } elseif ($activeDiff -lt 0) { "$activeDiff" } else { "no change" }
    $activeDiffClass = if ($activeDiff -gt 0) { 'bad' } elseif ($activeDiff -lt 0) { 'good' } else { 'na' }
    $activeDiffArrow = if ($activeDiff -gt 0) { 'up' } elseif ($activeDiff -lt 0) { 'down' } else { 'neutral' }
    $activeNote = "$prevActive as of $prevDate"
} else {
    $activeDiffFmt   = 'No prior data'
    $activeDiffClass = 'na'
    $activeDiffArrow = 'neutral'
    $activeNote      = 'First run — prior data will appear next time'
}
$closureRate = [math]::Round($closed / $ytdTotal * 100, 1)

# SBA (combined)
$sba = ($cases | Where-Object { $_.L1_CATEGORY -match 'System Based' }).Count
$sbaPct = [math]::Round($sba / $ytdTotal * 100, 1)

# Monthly volume
$months     = @('January','February','March','April','May','June','July','August','September','October','November','December')
$monthDays  = @{January=31;February=28;March=31;April=30;May=31;June=30;July=31;August=31;September=30;October=31;November=30;December=31}

$monthlyData = @()
foreach ($m in $months) {
    $mRows = $cases | Where-Object { $_.MONTH -eq $m }
    if ($mRows.Count -eq 0) { continue }
    $mSB    = ($mRows | Where-Object { $_.L1_CATEGORY -match 'System Based' }).Count
    $mPS    = ($mRows | Where-Object { $_.L1_CATEGORY -match 'Practice System' }).Count
    $mOther = $mRows.Count - $mSB - $mPS
    $mClosed = ($mRows | Where-Object { $_.STATUS -in @('Closed','Completed') }).Count
    $mClosureRate = [math]::Round($mClosed / $mRows.Count * 100, 1)
    $days   = if ($m -eq $CurrentMonth) { ($today - [datetime]"Jan 1, $($today.Year)").Days - ($months[0..($months.IndexOf($m)-1)] | ForEach-Object { $monthDays[$_] } | Measure-Object -Sum).Sum } else { $monthDays[$m] }
    # simpler: just use actual day count
    $daysInMonth = switch ($m) {
        'January'   { 31 } 'February' { 28 } 'March'    { 31 } 'April'    { 30 }
        'May'       { 31 } 'June'     { 30 } 'July'     { 31 } 'August'   { 31 }
        'September' { 30 } 'October'  { 31 } 'November' { 30 } 'December' { 31 }
    }
    $isMtd   = $m -eq $CurrentMonth
    $dayCount = if ($isMtd) { $today.Day } else { $daysInMonth }
    $avgDay  = [math]::Round($mRows.Count / $dayCount, 1)
    $sbPct   = [math]::Round($mSB  / $mRows.Count * 100)
    $psPct   = [math]::Round($mPS  / $mRows.Count * 100)
    $otPct   = [math]::Round($mOther / $mRows.Count * 100)
    $monthlyData += [PSCustomObject]@{
        Month=$m; Short=$m.Substring(0,3); Total=$mRows.Count
        SB=$mSB; PS=$mPS; Other=$mOther
        SBPct=$sbPct; PSPct=$psPct; OtPct=$otPct
        AvgDay=$avgDay; IsMtd=$isMtd
        Closed=$mClosed; ClosureRate=$mClosureRate
    }
}

# Max total for bar scaling
$maxMonthTotal = ($monthlyData | Measure-Object -Property Total -Maximum).Maximum

# Active case aging
$activeCases = $cases | Where-Object { $_.STATUS -notin @('Closed','Completed') }
$age0_15  = ($activeCases | Where-Object { ($today - [datetime]::Parse($_.CREATED_ON_DATE)).Days -le 15 }).Count
$age16_30 = ($activeCases | Where-Object { $d=($today-[datetime]::Parse($_.CREATED_ON_DATE)).Days; $d -ge 16 -and $d -le 30 }).Count
$age30p   = ($activeCases | Where-Object { ($today - [datetime]::Parse($_.CREATED_ON_DATE)).Days -gt 30 }).Count
$age0pct  = [math]::Round($age0_15  / $active * 100, 1)
$age16pct = [math]::Round($age16_30 / $active * 100, 1)
$age30pct = [math]::Round($age30p   / $active * 100, 1)

# Top 5 courses (base code, strip version suffix)
$top5 = $cases | Where-Object { $_.EXT_COURSECODE -notin @('','None','N/A') } | ForEach-Object {
    [PSCustomObject]@{ Base = ($_.EXT_COURSECODE -replace '_\d{4}$','') }
} | Group-Object Base | Sort-Object Count -Descending | Select-Object -First 5
$maxCourse = $top5[0].Count

# Course descriptions lookup
$courseDesc = @{
    'C_AIG'    = 'SAP Certified - SAP Generative AI Developer'
    'C_ADBTP'  = 'SAP Certified - SAP BTP Administrator'
    'C_TS452'  = 'SAP Certified - S/4HANA Sourcing and Procurement'
    'C_ABAPD'  = 'SAP Certified - Associate ABAP Developer'
    'C_TS4FI'  = 'SAP Certified - S/4HANA Financial Accounting'
    'C_TS462'  = 'SAP Certified - S/4HANA Sales'
    'C_IEE2E'  = 'SAP Certified - SAP Integration Suite'
    'C_S4CFI'  = 'SAP Certified - S/4HANA Cloud Finance'
    'C_S4EWM'  = 'SAP Certified - S/4HANA Extended Warehouse Mgmt'
    'C_THR81'  = 'SAP Certified - SAP SuccessFactors Employee Central'
}

# ?? Load surveys ?????????????????????????????????????????????
$sHdr = @('DISPLAYID','CASETYPE','CASETYPEDESCRIPTION','MONTH','EXT_COURSECODE','EXT_MISCINFO',
          'SUBJECT','PRIORITY','SOURCE','SERVICETEAMNAME','ACCOUNTNAME','COMPLETED_ON','RATING',
          'COMMENT1','COMMENT2','RECORDEDDATE','PROCESSOR_ID','PROCESSOR_NAME','L1_CATEGORY',
          'L2_CATEGORY','DOMAIN_CATEGORY','FORWARDED_TO','SID_COUNT')
$surveys = Import-Csv $SurveysCsv -Header $sHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing CSAT metrics..." -ForegroundColor Yellow

$top2Ratings = @('Very Satisfied','Somewhat Satisfied')

# YTD CSAT
$ytdSurveys = $surveys.Count
$ytdT2B     = ($surveys | Where-Object { $_.RATING -in $top2Ratings }).Count
$ytdCSAT    = [math]::Round($ytdT2B / $ytdSurveys * 100, 2)
$csatDelta  = [math]::Round($ytdCSAT - $PY_CSAT_PCT, 2)

# Monthly CSAT
$csatMonthly = @()
foreach ($m in $months) {
    $mS = $surveys | Where-Object { $_.MONTH -eq $m }
    if ($mS.Count -eq 0) { continue }
    $t2b  = ($mS | Where-Object { $_.RATING -in $top2Ratings }).Count
    $pct  = [math]::Round($t2b / $mS.Count * 100, 2)
    $csatMonthly += [PSCustomObject]@{ Month=$m; Short=$m.Substring(0,3); N=$mS.Count; T2B=$t2b; Pct=$pct; IsMtd=($m -eq $CurrentMonth) }
}
$minCsat  = ($csatMonthly | Measure-Object -Property Pct -Minimum).Minimum
$lowestMo = ($csatMonthly | Sort-Object Pct | Select-Object -First 1).Month

# DSAT drivers for current month - static analysis (update manually each month if needed)
$dsatRows = $surveys | Where-Object { $_.MONTH -eq $CurrentMonth -and $_.RATING -in @('Very Dissatisfied','Somewhat Dissatisfied') }
$dsatCount = $dsatRows.Count

# SBA-specific CSAT
$sbaSurveys = $surveys | Where-Object { $_.L1_CATEGORY -match 'System Based' }
$sbaT2B     = ($sbaSurveys | Where-Object { $_.RATING -in $top2Ratings }).Count
$sbaCSAT    = [math]::Round($sbaT2B / $sbaSurveys.Count * 100, 2)
$sbaVsOverall = [math]::Round($sbaCSAT - $ytdCSAT, 2)
$sbaVsLabel = if ($sbaVsOverall -ge 0) { "$sbaVsOverall pts above overall avg ($ytdCSAT%)" } else { "$([math]::Abs($sbaVsOverall)) pts below overall avg ($ytdCSAT%)" }

# ?? Delta helpers ?????????????????????????????????????????????
function Get-Delta($current, $prior) {
    $d = [math]::Round($current - $prior)
    $pct = [math]::Round(($current - $prior) / $prior * 100)
    if ($pct -ge 0) { return "+$pct%" } else { return "$pct%" }
}
function Get-DeltaClass($current, $prior) {
    if ($current -ge $prior) { return 'up' } else { return 'down' }
}
function Get-ArrowClass($current, $prior) {
    if ($current -ge $prior) { return 'up' } else { return 'down' }
}

$volDelta      = Get-Delta $ytdTotal $PY_Volume
$volDeltaClass = Get-DeltaClass $ytdTotal $PY_Volume
$volArrow      = Get-ArrowClass $ytdTotal $PY_Volume

$sbaDelta      = Get-Delta $sba $PY_SBA
$sbaDeltaClass = Get-DeltaClass $sba $PY_SBA
$sbaArrow      = Get-ArrowClass $sba $PY_SBA

$csatDeltaFmt   = if ($csatDelta -ge 0) { "+$csatDelta pts" } else { "$csatDelta pts" }
$csatDeltaClass = if ($csatDelta -ge 0) { 'up' } else { 'down' }
$csatArrow      = if ($csatDelta -ge 0) { 'up' } else { 'down' }

# ?? Build monthly bar chart rows ?????????????????????????????
$barRows = ""
foreach ($m in $monthlyData) {
    $widthPct = [math]::Round($m.Total / $maxMonthTotal * 100, 1)
    $sbW  = [math]::Round($m.SB    / $maxMonthTotal * 100, 1)
    $psW  = [math]::Round($m.PS    / $maxMonthTotal * 100, 1)
    $otW  = [math]::Round($m.Other / $maxMonthTotal * 100, 1)
    $totalFmt = $m.Total.ToString('N0')
    $mtdTag = if ($m.IsMtd) { " <span style='font-size:10px;'>MTD</span>" } else { "" }
    $avgColor = if ($m.IsMtd) { "color:#F0AB00;" } else { "" }
    $totalClass = if ($m.IsMtd) { "chart-total mtd" } else { "chart-total" }
    $barRows += @"
        <div class="chart-row">
          <div class="chart-month">$($m.Short)</div>
          <div class="chart-bar-wrap">
            <div class="seg sb"    style="width:$($sbW)%"><span class="seg-label">$($m.SBPct)%</span></div>
            <div class="seg ps"    style="width:$($psW)%"><span class="seg-label">$($m.PSPct)%</span></div>
            <div class="seg other" style="width:$($otW)%">$(if($m.OtPct -ge 3){"<span class='seg-label'>$($m.OtPct)%</span>"})</div>
          </div>
          <div class="$totalClass">$totalFmt$mtdTag</div>
          <div class="chart-avg" style="$avgColor">$($m.AvgDay) <span>/ day</span></div>
        </div>
"@
}

# ?? Build monthly closure table rows ?????????????????????????
$closureRows = ""
foreach ($m in $monthlyData) {
    $pillClass = if ($m.ClosureRate -ge 95) { 'good' } elseif ($m.ClosureRate -ge 90) { 'warn' } else { 'bad' }
    $mtdLabel  = if ($m.IsMtd) { "$($m.Short) MTD" } else { $m.Month }
    $closureRows += @"
            <tr>
              <td>$mtdLabel</td>
              <td>$($m.Total.ToString('N0'))</td>
              <td>$($m.Closed.ToString('N0'))</td>
              <td><span class="rate-pill $pillClass">$($m.ClosureRate)%</span></td>
            </tr>
"@
}

# ?? ILT cases ????????????????????????????????????????????????
$iltCases   = $cases | Where-Object { $_.SUBJECT -match 'ILT' -or $_.EXT_MISCINFO -match 'ILT' }
$iltTotal      = $iltCases.Count
$iltPctOfTotal = [math]::Round($iltTotal / $ytdTotal * 100, 1)
$iltCdcCases   = $iltCases | Where-Object { $_.SUBJECT -match 'CDC' -or $_.EXT_MISCINFO -match 'CDC' }
$iltCdcCount   = $iltCdcCases.Count
$iltCdcPct     = [math]::Round($iltCdcCount / $iltTotal * 100, 1)

# ILT monthly
$iltMonthly = @()
foreach ($m in $months) {
    $mRows = $iltCases | Where-Object { $_.MONTH -eq $m }
    if ($mRows.Count -eq 0) { continue }
    $isMtd = $m -eq $CurrentMonth
    $daysInMonth = switch ($m) {
        'January'{31}'February'{28}'March'{31}'April'{30}'May'{31}'June'{30}
        'July'{31}'August'{31}'September'{30}'October'{31}'November'{30}'December'{31}
    }
    $dayCount = if ($isMtd) { $today.Day } else { $daysInMonth }
    $avgDay   = [math]::Round($mRows.Count / $dayCount, 1)
    $cdcCount = ($iltCdcCases | Where-Object { $_.MONTH -eq $m }).Count
    $iltMonthly += [PSCustomObject]@{ Month=$m; Short=$m.Substring(0,3); Total=$mRows.Count; CDC=$cdcCount; AvgDay=$avgDay; IsMtd=$isMtd }
}
$iltMaxMonth = ($iltMonthly | Measure-Object Total -Maximum).Maximum

# ILT monthly bar rows
$iltBarRows = ""
foreach ($m in $iltMonthly) {
    $w        = [math]::Round($m.Total / $iltMaxMonth * 100, 1)
    $mtdTag   = if ($m.IsMtd) { " <span style='font-size:10px;'>MTD</span>" } else { "" }
    $avgColor = if ($m.IsMtd) { "color:#F0AB00;" } else { "" }
    $totalClass = if ($m.IsMtd) { "chart-total mtd" } else { "chart-total" }
    $iltBarRows += @"
        <div class="chart-row">
          <div class="chart-month">$($m.Short)</div>
          <div class="chart-bar-wrap">
            <div class="seg" style="width:$($w)%;background:#38b6e8;"><span class="seg-label"></span></div>
          </div>
          <div class="$totalClass">$($m.Total)$mtdTag</div>
          <div class="chart-avg" style="$avgColor">$($m.AvgDay) <span>/ day</span></div>
        </div>
"@
}

# ILT top courses
$iltTop5 = $iltCases | Where-Object { $_.EXT_COURSECODE -notin @('','None','N/A') } | ForEach-Object {
    [PSCustomObject]@{ Base = ($_.EXT_COURSECODE -replace '_\d{4}$','') }
} | Group-Object Base | Sort-Object Count -Descending | Select-Object -First 5
$iltMaxCourse = if ($iltTop5) { ($iltTop5 | Measure-Object Count -Maximum).Maximum } else { 1 }

$iltCourseHtml = ""
$iltCourseColors = @('#38b6e8','#38b6e8','#F0AB00','#F0AB00','#2e5f7a')
$ci3 = 0
foreach ($c in $iltTop5) {
    $pct  = [math]::Round($c.Count / $iltTotal * 100)
    $w    = [math]::Round($c.Count / $iltMaxCourse * 100, 1)
    $desc = if ($courseDesc.ContainsKey($c.Name)) { $courseDesc[$c.Name] } else { $c.Name }
    $col  = $iltCourseColors[$ci3]
    $mb   = if ($ci3 -eq ($iltTop5.Count-1)) { "4px" } else { "14px" }
    $iltCourseHtml += @"
        <div style="margin-bottom:$mb;">
          <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;">
            <span style="font-size:13px;font-weight:700;color:#c9d1d9;">$($c.Name)</span>
            <span style="font-size:12px;font-weight:700;color:#38b6e8;">$($c.Count.ToString('N0')) <span style="font-size:10px;color:#4d606e;font-weight:400;">$pct%</span></span>
          </div>
          <div style="background:#21262d;border-radius:4px;height:8px;overflow:hidden;">
            <div style="width:$($w)%;height:100%;background:$col;border-radius:4px;"></div>
          </div>
        </div>
"@
    $ci3++
}

# ?? Build Top 5 courses ???????????????????????????????????????
$courseColors = @('#38b6e8','#38b6e8','#F0AB00','#F0AB00','#2e5f7a')
$courseHtml = ""
$ci2 = 0
foreach ($c in $top5) {
    $pct  = [math]::Round($c.Count / $ytdTotal * 100)
    $w    = [math]::Round($c.Count / $maxCourse * 100, 1)
    $desc = if ($courseDesc.ContainsKey($c.Name)) { $courseDesc[$c.Name] } else { $c.Name }
    $col  = $courseColors[$ci2]
    $isLast = $ci2 -eq ($top5.Count - 1)
    $mb = if ($isLast) { "4px" } else { "14px" }
    $courseHtml += @"
        <div style="margin-bottom:$mb;">
          <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;">
            <span style="font-size:13px;font-weight:700;color:#c9d1d9;">$($c.Name)</span>
            <span style="font-size:12px;font-weight:700;color:#38b6e8;">$($c.Count.ToString('N0')) <span style="font-size:10px;color:#4d606e;font-weight:400;">$pct%</span></span>
          </div>
          <div style="font-size:10px;color:#4d606e;margin-bottom:5px;">$desc</div>
          <div style="background:#21262d;border-radius:4px;height:8px;overflow:hidden;">
            <div style="width:$($w)%;height:100%;background:$col;border-radius:4px;"></div>
          </div>
        </div>
"@
    $ci2++
}

# ?? Key highlights (auto-generated) ??????????????????????????
$junClosureRow = $monthlyData | Where-Object { $_.IsMtd }
$junClosure    = if ($junClosureRow) { $junClosureRow.ClosureRate } else { 0 }
$csatLowestPct = $minCsat
$csatLatest    = ($csatMonthly | Select-Object -Last 1).Pct

$highlight1 = "<strong>SBA volume up $sbaDelta YoY</strong> - System Based Assessment cases surged from $($PY_SBA.ToString('N0')) (2025 FY) to $($sba.ToString('N0')) YTD, driven by certification sprint and expanded program delivery."
$highlight2 = "<strong>CSAT at $ytdCSAT% vs $CSAT_TARGET% target</strong> - YTD score $csatDeltaFmt vs 2025. $($CurrentMonthShort) MTD at $($csatLatest)% - $(if($csatLatest -ge $CSAT_TARGET){'above target check'}else{'trending up since ' + $lowestMo + ' low of ' + $csatLowestPct + '%'})."
$highlight3 = "<strong>$($CurrentMonthShort) closure rate at $junClosure%</strong> - $(if($junClosure -lt 90){'Reflects high mid-month inflow. Active backlog of ' + $active.ToString('N0') + ' cases requires monitoring.'}else{'Strong closure performance maintained. Active backlog: ' + $active.ToString('N0') + ' cases.'})"

# ?? Write HTML ????????????????????????????????????????????????
$outputFile = "$OutputDir\GTS_Executive_Summary.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>GTS ESM - Executive Summary YTD 2026</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #0d1117; color: #e6edf3; padding: 36px 32px; max-width: 1100px; margin: 0 auto; }
    .page-header { margin-bottom: 32px; padding-bottom: 18px; border-bottom: 3px solid #007DB8; display: flex; justify-content: space-between; align-items: flex-end; }
    .page-header h1 { font-size: 24px; font-weight: 700; color: #38b6e8; letter-spacing: 0.2px; }
    .page-header .subtitle { font-size: 13px; color: #7d8fa1; margin-top: 4px; }
    .page-header .datestamp { font-size: 12px; color: #4d606e; text-align: right; }
    .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 36px; }
    .kpi-card { background: #161b22; border-radius: 10px; padding: 20px 18px 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.4); border-left: 4px solid var(--accent); position: relative; }
    .kpi-card.blue   { --accent: #38b6e8; } .kpi-card.gold  { --accent: #F0AB00; }
    .kpi-card.blue2  { --accent: #007DB8; } .kpi-card.gold2 { --accent: #C88A00; }
    .kpi-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #7d8fa1; margin-bottom: 8px; }
    .kpi-value { font-size: 34px; font-weight: 700; color: var(--accent); line-height: 1; }
    .kpi-footer { display: flex; align-items: center; justify-content: space-between; margin-top: 10px; }
    .kpi-sub { font-size: 11px; color: #4d606e; }
    .kpi-delta { display: flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 700; padding: 3px 8px; border-radius: 20px; }
    .kpi-delta.up   { background: #0d2818; color: #3fb950; } .kpi-delta.down { background: #2d0f0f; color: #f85149; } .kpi-delta.na { background: #1c2128; color: #4d606e; }
    .arrow { display: inline-block; width: 0; height: 0; flex-shrink: 0; }
    .arrow.up    { border-left: 5px solid transparent; border-right: 5px solid transparent; border-bottom: 8px solid #3fb950; }
    .arrow.down  { border-left: 5px solid transparent; border-right: 5px solid transparent; border-top: 8px solid #f85149; }
    .arrow.neutral { border-left: 5px solid transparent; border-right: 5px solid transparent; border-bottom: 8px solid #9ca3af; opacity: 0.4; }
    .tooltip-icon { position: absolute; top: 12px; right: 12px; width: 16px; height: 16px; border-radius: 50%; background: #21262d; color: #7d8fa1; font-size: 10px; font-weight: 700; display: flex; align-items: center; justify-content: center; cursor: default; user-select: none; }
    .tooltip-icon:hover .tooltip-box { display: block; }
    .tooltip-box { display: none; position: absolute; top: 22px; right: 0; width: 200px; background: #21262d; color: #c9d1d9; font-size: 11px; font-weight: 400; line-height: 1.5; padding: 10px 12px; border-radius: 8px; border: 1px solid #30363d; box-shadow: 0 8px 24px rgba(0,0,0,0.5); z-index: 10; pointer-events: none; }
    .tooltip-box::before { content: ''; position: absolute; top: -5px; right: 4px; border-left: 5px solid transparent; border-right: 5px solid transparent; border-bottom: 5px solid #30363d; }
    .highlights { background: #161b22; border-radius: 10px; padding: 18px 22px; margin-bottom: 36px; border-left: 4px solid #F0AB00; box-shadow: 0 2px 8px rgba(0,0,0,0.4); }
    .highlights-title { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #F0AB00; margin-bottom: 12px; }
    .highlight-item { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 8px; font-size: 12px; color: #c9d1d9; line-height: 1.5; }
    .highlight-item:last-child { margin-bottom: 0; }
    .hi-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; margin-top: 5px; }
    .hi-dot.win { background: #3fb950; } .hi-dot.risk { background: #f85149; } .hi-dot.trend { background: #38b6e8; }
    .section { margin-bottom: 36px; }
    .section-header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; }
    .section-header h2 { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; color: #38b6e8; white-space: nowrap; }
    .section-rule { flex: 1; height: 1px; background: linear-gradient(to right, #F0AB00, transparent); }
    .chart-wrap { background: #161b22; border-radius: 10px; padding: 24px 24px 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.4); }
    .chart-row { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; font-size: 12px; }
    .chart-month { width: 36px; font-size: 12px; font-weight: 600; color: #7d8fa1; flex-shrink: 0; text-align: right; }
    .chart-bar-wrap { flex: 1; height: 34px; display: flex; border-radius: 4px; overflow: hidden; background: #21262d; }
    .seg { height: 100%; transition: width 0.3s; display: flex; align-items: center; justify-content: center; overflow: hidden; }
    .seg-label { font-size: 10px; font-weight: 700; color: rgba(255,255,255,0.9); white-space: nowrap; padding: 0 4px; }
    .seg.sb { background: #38b6e8; } .seg.ps { background: #F0AB00; } .seg.other { background: #2e5f7a; }
    .chart-total { width: 58px; font-size: 12px; font-weight: 700; color: #c9d1d9; text-align: right; flex-shrink: 0; }
    .chart-total.mtd { color: #7d8fa1; }
    .chart-avg { width: 72px; font-size: 11px; font-weight: 500; color: #7d8fa1; text-align: right; flex-shrink: 0; }
    .chart-avg span { font-size: 10px; color: #4d606e; }
    .chart-legend { display: flex; gap: 20px; margin-top: 16px; padding-top: 12px; border-top: 1px solid #21262d; }
    .legend-item { display: flex; align-items: center; gap: 6px; font-size: 11px; color: #7d8fa1; }
    .legend-dot { width: 12px; height: 12px; border-radius: 2px; flex-shrink: 0; }
    .ops-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
    .ops-stat-row { display: flex; align-items: center; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #21262d; font-size: 13px; }
    .ops-stat-row:last-child { border-bottom: none; }
    .ops-stat-label { color: #7d8fa1; }
    .ops-stat-val { font-weight: 700; color: #c9d1d9; font-size: 14px; }
    .ops-stat-val.good { color: #3fb950; } .ops-stat-val.warn { color: #F0AB00; } .ops-stat-val.bad { color: #f85149; }
    .aging-row { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; font-size: 12px; }
    .aging-label { width: 60px; color: #7d8fa1; flex-shrink: 0; }
    .aging-track { flex: 1; background: #21262d; border-radius: 4px; height: 20px; overflow: hidden; display: flex; align-items: center; }
    .aging-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 8px; }
    .aging-fill span { font-size: 10px; font-weight: 700; color: rgba(255,255,255,0.9); }
    .aging-count { width: 36px; text-align: right; font-weight: 700; color: #c9d1d9; font-size: 12px; }
    .closure-table { width: 100%; border-collapse: collapse; margin-top: 4px; }
    .closure-table th { font-size: 10px; font-weight: 600; color: #4d606e; text-transform: uppercase; letter-spacing: 0.5px; padding: 6px 8px; text-align: left; border-bottom: 1px solid #21262d; }
    .closure-table td { font-size: 12px; padding: 7px 8px; border-bottom: 1px solid #21262d; color: #c9d1d9; }
    .closure-table tr:last-child td { border-bottom: none; }
    .rate-pill { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 700; }
    .rate-pill.good { background: #0d2818; color: #3fb950; } .rate-pill.warn { background: #2a2000; color: #F0AB00; } .rate-pill.bad { background: #2d0f0f; color: #f85149; }
    .page-footer { margin-top: 12px; font-size: 11px; color: #30363d; text-align: center; padding-top: 16px; border-top: 1px solid #21262d; }
  </style>
</head>
<body>

  <div class="page-header">
    <div>
      <h1>GTS ESM - Operations Executive Summary</h1>
      <div class="subtitle">Global Training Support &nbsp;.&nbsp; Year-to-Date 2026</div>
    </div>
    <div class="datestamp">As of $ReportDate<br>Source: SERVICE_CLOUD</div>
  </div>

  <div class="kpi-row">
    <div class="kpi-card blue">
      <div class="tooltip-icon">i<div class="tooltip-box">YTD $YtdLabel. Full year 2025 was $($PY_Volume.ToString('N0')) cases. Volume growth driven by the System Based Assessment program ramp-up.</div></div>
      <div class="kpi-label">YTD Case Volume</div>
      <div class="kpi-value">$($ytdTotal.ToString('N0'))</div>
      <div class="kpi-footer">
        <span class="kpi-sub">vs 2025 FY: $($PY_Volume.ToString('N0'))</span>
        <span class="kpi-delta $volDeltaClass"><span class="arrow $volArrow"></span> $volDelta</span>
      </div>
    </div>
    <div class="kpi-card gold2">
      <div class="tooltip-icon">i<div class="tooltip-box">Combines System Based Assessment and Exam - System Based Assessment categories. 2025 full year was $($PY_SBA.ToString('N0')) cases.</div></div>
      <div class="kpi-label">System Based Cases</div>
      <div class="kpi-value">$($sba.ToString('N0'))</div>
      <div class="kpi-footer">
        <span class="kpi-sub">$sbaPct% of total volume &nbsp;·&nbsp; vs 2025 FY: $($PY_SBA.ToString('N0'))</span>
        <span class="kpi-delta $sbaDeltaClass"><span class="arrow $sbaArrow"></span> $sbaDelta</span>
      </div>
    </div>
    <div class="kpi-card blue2">
      <div class="tooltip-icon">i<div class="tooltip-box">Point-in-time snapshot as of $ReportDate. Includes all cases with status: In Process, Customer Action, Awaiting Third Party, and Open.</div></div>
      <div class="kpi-label">Active Cases</div>
      <div class="kpi-value">$($active.ToString('N0'))</div>
      <div class="kpi-footer">
        <span class="kpi-sub">$activeNote</span>
        <span class="kpi-delta $activeDiffClass"><span class="arrow $activeDiffArrow"></span> $activeDiffFmt</span>
      </div>
    </div>
    <div class="kpi-card gold">
      <div class="tooltip-icon">i<div class="tooltip-box">2026 YTD CSAT based on $($ytdSurveys.ToString('N0')) survey responses. Top-2-Box = Very Satisfied + Somewhat Satisfied. 2025 full year was $PY_CSAT_PCT% from $($PY_CSAT_N.ToString('N0')) responses.</div></div>
      <div class="kpi-label">CSAT Score</div>
      <div class="kpi-value">$ytdCSAT%</div>
      <div class="kpi-footer">
        <span class="kpi-sub">$($ytdSurveys.ToString('N0')) surveys &nbsp;·&nbsp; vs 2025 FY: $PY_CSAT_PCT% ($($PY_CSAT_N.ToString('N0')) surveys)</span>
        <span class="kpi-delta $csatDeltaClass"><span class="arrow $csatArrow"></span> $csatDeltaFmt</span>
      </div>
    </div>
  </div>

  <div class="highlights">
    <div class="highlights-title"> Key Highlights</div>
    <div class="highlight-item"><div class="hi-dot win"></div><div>$highlight1</div></div>
    <div class="highlight-item"><div class="hi-dot risk"></div><div>$highlight2</div></div>
    <div class="highlight-item"><div class="hi-dot trend"></div><div>$highlight3</div></div>
  </div>

  <div class="section">
    <div class="section-header"><h2>Monthly Case Volume</h2><div class="section-rule"></div></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
      <div class="chart-wrap">
        <div class="chart-row" style="margin-bottom:6px;">
          <div class="chart-month" style="color:#4d606e;font-size:10px;">Month</div>
          <div style="flex:1;"></div>
          <div class="chart-total" style="font-size:10px;color:#4d606e;font-weight:600;">Total</div>
          <div class="chart-avg" style="font-size:10px;color:#4d606e;font-weight:600;">Avg/Day</div>
        </div>
        $barRows
        <div class="chart-legend">
          <div class="legend-item"><div class="legend-dot" style="background:#38b6e8"></div> System Based Assessment</div>
          <div class="legend-item"><div class="legend-dot" style="background:#F0AB00"></div> Practice Systems</div>
          <div class="legend-item"><div class="legend-dot" style="background:#2e5f7a"></div> Other</div>
        </div>
      </div>
      <div class="chart-wrap">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:16px;">Top 5 Courses by Case Volume (YTD)</div>
        $courseHtml
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-header"><h2>Instructor-Led Training (ILT) Case Volume</h2><div class="section-rule"></div></div>
    <div style="display:grid;grid-template-columns:160px 1fr 1fr;gap:20px;align-items:start;">
      <!-- YTD square -->
      <div class="chart-wrap" style="text-align:center;padding:24px 16px;">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:12px;">Instructor-Led Training (ILT) YTD</div>
        <div style="font-size:42px;font-weight:700;color:#38b6e8;line-height:1;">$($iltTotal.ToString('N0'))</div>
        <div style="font-size:11px;color:#4d606e;margin-top:8px;">$iltPctOfTotal% of total cases</div>
        <div style="margin-top:14px;padding-top:12px;border-top:1px solid #21262d;">
          <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:6px;">CDC</div>
          <div style="font-size:24px;font-weight:700;color:#F0AB00;line-height:1;">$($iltCdcCount.ToString('N0'))</div>
          <div style="font-size:11px;color:#4d606e;margin-top:4px;">$iltCdcPct% of ILT cases</div>
        </div>
      </div>
      <!-- Monthly trend -->
      <div class="chart-wrap">
        <div class="chart-row" style="margin-bottom:6px;">
          <div class="chart-month" style="color:#4d606e;font-size:10px;">Month</div>
          <div style="flex:1;"></div>
          <div class="chart-total" style="font-size:10px;color:#4d606e;font-weight:600;">Total</div>
          <div class="chart-avg" style="font-size:10px;color:#4d606e;font-weight:600;">Avg/Day</div>
        </div>
        $iltBarRows
      </div>
      <!-- Top courses -->
      <div class="chart-wrap">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:16px;">Top 5 ILT Courses (YTD)</div>
        $iltCourseHtml
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-header"><h2>Operational Health</h2><div class="section-rule"></div></div>
    <div class="ops-grid">
      <div class="chart-wrap">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:12px;">Case Health Indicators</div>
        <div class="ops-stat-row">
          <span class="ops-stat-label">YTD Closure Rate</span>
          <span class="ops-stat-val $(if($closureRate -ge 90){'good'}elseif($closureRate -ge 80){'warn'}else{'bad'})">$closureRate%</span>
        </div>
        <div class="ops-stat-row">
          <span class="ops-stat-label">Active Cases</span>
          <span class="ops-stat-val warn">$($active.ToString('N0'))</span>
        </div>
        <div class="ops-stat-row">
          <span class="ops-stat-label">Cases &gt;30 Days</span>
          <span class="ops-stat-val $(if($age30p -eq 0){'good'}elseif($age30p -le 50){'warn'}else{'bad'})">$($age30p.ToString('N0'))</span>
        </div>
        <div style="margin-top:18px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:10px;">Active Case Aging</div>
        <div class="aging-row">
          <div class="aging-label">0 - 15d</div>
          <div class="aging-track"><div class="aging-fill" style="width:$($age0pct)%;background:#3fb950;"><span>$age0pct%</span></div></div>
          <div class="aging-count">$($age0_15.ToString('N0'))</div>
        </div>
        <div class="aging-row">
          <div class="aging-label">16 - 30d</div>
          <div class="aging-track"><div class="aging-fill" style="width:$($age16pct)%;background:#F0AB00;"><span>$age16pct%</span></div></div>
          <div class="aging-count">$($age16_30.ToString('N0'))</div>
        </div>
        <div class="aging-row">
          <div class="aging-label">&gt; 30d</div>
          <div class="aging-track"><div class="aging-fill" style="width:$($age30pct)%;background:#f85149;"><span>$age30pct%</span></div></div>
          <div class="aging-count">$($age30p.ToString('N0'))</div>
        </div>
      </div>
      <div class="chart-wrap">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7d8fa1;margin-bottom:12px;">Monthly Closure Rate</div>
        <table class="closure-table">
          <thead><tr><th>Month</th><th>Received</th><th>Closed</th><th>Rate</th></tr></thead>
          <tbody>$closureRows</tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="page-footer">Generated $ReportDate &nbsp;.&nbsp; GTS ESM Operations &nbsp;.&nbsp; Cases: $(Split-Path $CasesXlsx -Leaf) &nbsp;.&nbsp; Surveys: $(Split-Path $SurveysXlsx -Leaf)</div>

</body>
</html>
"@

$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
