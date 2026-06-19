# ============================================================
# GTS ESM CSAT Bulletin - Auto-Generator
#
# Usage:
#   1. Drop updated xlsx files in Downloads
#   2. Run: powershell -ExecutionPolicy Bypass -File "Generate_CSAT_Bulletin.ps1"
# ============================================================

# ?? AUTO-DETECT FILES ????????????????????????????????????????
$SurveysXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Surveys_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$CasesXlsx   = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx"   | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$OutputDir   = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"

# Derive report date from surveys filename
$baseName   = [System.IO.Path]::GetFileNameWithoutExtension($SurveysXlsx)
$datePart   = $baseName -replace '^GTS ESM Surveys_', ''
$parsedDate = $null
foreach ($fmt in @('MMMMd','MMMMdd','MMMd','MMMdd')) {
    try { $parsedDate = [datetime]::ParseExact($datePart, $fmt, [System.Globalization.CultureInfo]::InvariantCulture); break } catch {}
}
if (-not $parsedDate) { $parsedDate = [datetime]::Today }
$parsedDate = $parsedDate.AddYears([datetime]::Today.Year - $parsedDate.Year)

$ReportDate        = $parsedDate.ToString("MMMM d, yyyy")
$ReportDateShort   = $parsedDate.ToString("MMM d")
$CurrentMonth      = $parsedDate.ToString("MMMM")
$CurrentMonthShort = $parsedDate.ToString("MMM")

# Targets
$CSAT_TARGET = 80

$SurveysCsv = "$env:TEMP\gts_surveys_csat.csv"
$CasesCsv   = "$env:TEMP\gts_cases_csat.csv"

# ?? EXPORT XLSX ? CSV ????????????????????????????????????????
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
Export-XlsxToCsv $SurveysXlsx $SurveysCsv
Export-XlsxToCsv $CasesXlsx   $CasesCsv

# ?? LOAD DATA ????????????????????????????????????????????????
$sHdr = @('DISPLAYID','CASETYPE','CASETYPEDESCRIPTION','MONTH','EXT_COURSECODE','EXT_MISCINFO',
          'SUBJECT','PRIORITY','SOURCE','SERVICETEAMNAME','ACCOUNTNAME','COMPLETED_ON','RATING',
          'COMMENT1','COMMENT2','RECORDEDDATE','PROCESSOR_ID','PROCESSOR_NAME','L1_CATEGORY',
          'L2_CATEGORY','DOMAIN_CATEGORY','FORWARDED_TO','SID_COUNT')
$surveys = Import-Csv $SurveysCsv -Header $sHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing CSAT metrics..." -ForegroundColor Yellow

$top2   = @('Very Satisfied','Somewhat Satisfied')
$dsatRatings = @('Very Dissatisfied','Somewhat Dissatisfied')
$months = @('January','February','March','April','May','June','July','August','September','October','November','December')

# ?? YTD KPIs ??????????????????????????????????????????????????
$ytdN    = $surveys.Count
$ytdT2B  = ($surveys | Where-Object { $_.RATING -in $top2 }).Count
$ytdCSAT = [math]::Round($ytdT2B / $ytdN * 100, 2)
$ytdDSAT = ($surveys | Where-Object { $_.RATING -in $dsatRatings }).Count

# MTD KPIs
$mtdSurveys = $surveys | Where-Object { $_.MONTH -eq $CurrentMonth }
$mtdN    = $mtdSurveys.Count
$mtdT2B  = ($mtdSurveys | Where-Object { $_.RATING -in $top2 }).Count
$mtdCSAT = if ($mtdN -gt 0) { [math]::Round($mtdT2B / $mtdN * 100, 2) } else { 0 }
$mtdDSAT = ($mtdSurveys | Where-Object { $_.RATING -in $dsatRatings }).Count

# CSAT vs target
$csatVsTarget  = [math]::Round($ytdCSAT - $CSAT_TARGET, 2)
$csatVsTargetFmt = if ($csatVsTarget -ge 0) { "+$csatVsTarget pts vs target" } else { "$csatVsTarget pts vs target" }

# ?? SBA vs PRACTICE SYSTEMS CSAT ????????????????????????????????
$sbaSurveys = $surveys | Where-Object { $_.L1_CATEGORY -match 'System Based' }
$sbaN       = $sbaSurveys.Count
$sbaT2B     = ($sbaSurveys | Where-Object { $_.RATING -in $top2 }).Count
$sbaCSAT    = if ($sbaN -gt 0) { [math]::Round($sbaT2B / $sbaN * 100, 2) } else { 0 }
$sbaDSAT    = ($sbaSurveys | Where-Object { $_.RATING -in $dsatRatings }).Count

$psSurveys  = $surveys | Where-Object { $_.L1_CATEGORY -match 'Practice System' }
$psN        = $psSurveys.Count
$psT2B      = ($psSurveys | Where-Object { $_.RATING -in $top2 }).Count
$psCSAT     = if ($psN -gt 0) { [math]::Round($psT2B / $psN * 100, 2) } else { 0 }
$psDSAT     = ($psSurveys | Where-Object { $_.RATING -in $dsatRatings }).Count

$sbaShare   = [math]::Round($sbaN / $ytdN * 100)
$psShare    = [math]::Round($psN  / $ytdN * 100)

# ILT surveys
$iltSurveys = $surveys | Where-Object { $_.SUBJECT -match 'ILT' -or $_.EXT_MISCINFO -match 'ILT' }
$iltN       = $iltSurveys.Count
$iltT2B     = ($iltSurveys | Where-Object { $_.RATING -in $top2 }).Count
$iltCSAT    = if ($iltN -gt 0) { [math]::Round($iltT2B / $iltN * 100, 2) } else { 0 }
$iltDSAT    = ($iltSurveys | Where-Object { $_.RATING -in $dsatRatings }).Count
$iltShare   = [math]::Round($iltN / $ytdN * 100)

# ?? MONTHLY TREND ???????????????????????????????????????????
$monthlyTrend = @()
foreach ($m in $months) {
    $mS = $surveys | Where-Object { $_.MONTH -eq $m }
    if ($mS.Count -eq 0) { continue }
    $mT2B  = ($mS | Where-Object { $_.RATING -in $top2 }).Count
    $mDSAT = ($mS | Where-Object { $_.RATING -in $dsatRatings }).Count
    $mPct  = [math]::Round($mT2B / $mS.Count * 100, 2)
    $monthlyTrend += [PSCustomObject]@{
        Month=$m; Short=$m.Substring(0,3); N=$mS.Count; T2B=$mT2B; DSAT=$mDSAT; Pct=$mPct; IsMtd=($m -eq $CurrentMonth)
    }
}
$prevMonth = $monthlyTrend | Where-Object { -not $_.IsMtd } | Select-Object -Last 1
$mtdVsPrev = if ($prevMonth) { [math]::Round($mtdCSAT - $prevMonth.Pct, 2) } else { 0 }
$mtdVsPrevFmt = if ($mtdVsPrev -ge 0) { "+$mtdVsPrev pts vs $($prevMonth.Short)" } else { "$mtdVsPrev pts vs $($prevMonth.Short)" }

# ?? DSAT DRIVERS (L1 category) ??????????????????????????????
$dsatAll = $surveys | Where-Object { $_.RATING -in $dsatRatings }

# Normalise: merge Exam - System Based into SBA; everything else -> Others
$dsatNormalised = $dsatAll | ForEach-Object {
    $cat = $_.L1_CATEGORY
    if     ($cat -match 'Exam - System Based' -or $cat -match '^System Based') { $cat = 'System Based Assessment' }
    elseif ($cat -match 'Practice System')                                      { $cat = 'Practice Systems' }
    else                                                                        { $cat = 'Others' }
    [PSCustomObject]@{ L1 = $cat }
}
$dsatByL1  = $dsatNormalised | Group-Object L1 | Sort-Object Count -Descending
$maxDsatL1 = ($dsatByL1 | Measure-Object Count -Maximum).Maximum

# ?? COURSES WITH MOST DSATs ??????????????????????????
$dsatCourses = $dsatAll | Where-Object { $_.EXT_COURSECODE -notin @('','None','N/A') } | ForEach-Object {
    [PSCustomObject]@{ Base = ($_.EXT_COURSECODE -replace '_\d{4}$','') }
} | Group-Object Base | Sort-Object Count -Descending | Select-Object -First 5
$maxDsatCourse = if ($dsatCourses) { ($dsatCourses | Measure-Object Count -Maximum).Maximum } else { 1 }

# ?? VOICE OF CUSTOMER - COMMENT BUCKETING (current month) ???????????
$mtdAllComments  = $surveys | Where-Object { $_.MONTH -eq $CurrentMonth -and ($_.COMMENT1.Trim() -ne '' -or $_.COMMENT2.Trim() -ne '') }
$mtdDsatComments = $mtdAllComments | Where-Object { $_.RATING -in $dsatRatings }
$mtdSatComments  = $mtdAllComments | Where-Object { $_.RATING -in $top2 }

# Negative buckets
$negBuckets = [ordered]@{
    'Issue Not Resolved'        = @{ keywords = @('not resolved','not solve','unresolved','still','same issue','issue not','did not resolve','wasn''t resolved','unable to solve','cannot solve','can''t solve','no fix','no solution','not fixed','not helped','did not help'); color = '#e05252' }
    'Exam / System Issues'      = @{ keywords = @('exam','timer','log off','logout','log out','reschedule','attempt','rescheduled','system error','technical issue','certification tool','3h','4h','access','instance','re-take','retake','retak','practice system'); color = '#8b5cf6' }
    'Agent Knowledge'           = @{ keywords = @('knowledge','don''t know','no knowledge','poor','not upto mark','upto mark','not able to understand','can''t understand','does not understand','not understand','incompetent','untrained','unqualified','working in sap','people who don'); color = '#f97316' }
    'Slow / Delayed Response'   = @{ keywords = @('slow','delay','late','days','hours','waiting','wait','too long','long time','took time','time taken','unacceptable','15 hour','10 day','demorad','demorada'); color = '#f59e0b' }
    'Case Closed Prematurely'   = @{ keywords = @('closed','premature','close the case','closed the ticket','case was closed','closed before','not complete','incomplete'); color = '#ec4899' }
    'Exam Result Dispute'       = @{ keywords = @('result','score','review','zero','mistake','saving','saved','submission','reevaluation','re-evaluation','wrong result','failed','fail','exam result'); color = '#64748b' }
}

# Positive buckets
$posBuckets = [ordered]@{
    'Quick Resolution'      = @{ keywords = @('quick','fast','prompt','rapid','immediate','timely','swift','soon','quickly','promptly'); color = '#22c55e' }
    'Problem Solved'        = @{ keywords = @('resolved','solved','fixed','solution','worked','working now','issue resolved','problem solved','sorted','helped me'); color = '#10b981' }
    'Helpful & Friendly'    = @{ keywords = @('helpful','friendly','kind','polite','great','excellent','wonderful','amazing','fantastic','professional','courteous','nice','good support','good service'); color = '#007DB8' }
    'Clear Communication'   = @{ keywords = @('clear','explained','understand','easy','simple','straightforward','well explained','good explanation','detailed','thorough'); color = '#38b6e8' }
}

function Get-BucketedRows($comments, $bucketDef) {
    $rows = [ordered]@{}
    foreach ($key in $bucketDef.Keys) { $rows[$key] = @() }
    $rows['Other'] = @()
    foreach ($r in $comments) {
        $text = if ($r.COMMENT1.Trim() -ne '') { $r.COMMENT1.Trim().ToLower() } else { $r.COMMENT2.Trim().ToLower() }
        $matched = $false
        foreach ($key in $bucketDef.Keys) {
            foreach ($kw in $bucketDef[$key].keywords) {
                if ($text -match [regex]::Escape($kw)) {
                    $rows[$key] += $r; $matched = $true; break
                }
            }
            if ($matched) { break }
        }
        if (-not $matched) { $rows['Other'] += $r }
    }
    return $rows
}

$negRows = Get-BucketedRows $mtdDsatComments $negBuckets
$posRows = Get-BucketedRows $mtdSatComments  $posBuckets

function Build-VocColumn($bucketDef, $bucketRows, $total, $otherColor, $insightMap) {
    $html = ""
    $allKeys = @($bucketDef.Keys) + @('Other')
    foreach ($key in $allKeys) {
        $rows = $bucketRows[$key]
        if ($rows.Count -eq 0) { continue }
        $col     = if ($key -eq 'Other') { $otherColor } else { $bucketDef[$key].color }
        $pct     = [math]::Round($rows.Count / $total * 100)
        $wPct    = [math]::Round($rows.Count / $total * 100, 1)
        $insight = if ($insightMap.ContainsKey($key)) { $insightMap[$key] } else { "Miscellaneous responses that did not match a specific theme." }
        $html += @"
      <div style="background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,0.06);border-left:4px solid $col;margin-bottom:10px;">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
          <span style="font-size:13px;font-weight:700;color:#1e293b;">$key</span>
          <span style="font-size:12px;font-weight:700;color:$col;">$($rows.Count) <span style="font-size:10px;color:#94a3b8;font-weight:400;">($pct%)</span></span>
        </div>
        <div style="background:#e2e8f0;border-radius:4px;height:7px;overflow:hidden;margin-bottom:8px;">
          <div style="width:$($wPct)%;height:100%;background:$col;border-radius:4px;opacity:0.75;"></div>
        </div>
        <div style="font-size:11px;color:#64748b;line-height:1.6;">$insight</div>
      </div>
"@
    }
    return $html
}

$negInsights = @{
    'Issue Not Resolved'        = "Top complaint in $CurrentMonthShort - $($negRows['Issue Not Resolved'].Count) customers report their problem remained unresolved at case closure. Repeat contacts and reopen risk are high."
    'Agent Knowledge'           = "$($negRows['Agent Knowledge'].Count) responses cite agents providing generic or incorrect replies. Query comprehension and technical depth need reinforcement."
    'Slow / Delayed Response'   = "$($negRows['Slow / Delayed Response'].Count) customers flagged long wait times. ILT and live exam issues are most time-sensitive and need priority handling."
    'Exam / System Issues'      = "$($negRows['Exam / System Issues'].Count) cases involve technical problems during live exams - timer confusion, unexpected logouts, and rescheduling friction."
    'Case Closed Prematurely'   = "$($negRows['Case Closed Prematurely'].Count) customers indicate their case was closed before the issue was fully addressed."
    'Exam Result Dispute'       = "$($negRows['Exam Result Dispute'].Count) learners contest exam outcomes citing data conflicts or submission errors with no review provided."
    'Other'                     = "$($negRows['Other'].Count) responses did not match a specific theme and may warrant manual review."
}

$posInsights = @{
    'Quick Resolution'      = "$($posRows['Quick Resolution'].Count) customers called out fast turnaround as a key strength - prompt handling is a visible differentiator for the team."
    'Helpful & Friendly'    = "$($posRows['Helpful & Friendly'].Count) responses praised agent attitude and professionalism - learners appreciated the supportive tone."
    'Clear Communication'   = "$($posRows['Clear Communication'].Count) customers found responses easy to understand and well-explained, reducing repeat contacts."
    'Problem Solved'        = "$($posRows['Problem Solved'].Count) customers confirmed their issue was fully resolved - a direct driver of satisfaction scores."
    'Other'                 = "$($posRows['Other'].Count) positive responses reflect general satisfaction not tied to a specific theme."
}

$negColHtml = Build-VocColumn $negBuckets $negRows $mtdDsatComments.Count '#94a3b8' $negInsights
$posColHtml = Build-VocColumn $posBuckets $posRows $mtdSatComments.Count  '#94a3b8' $posInsights
$vocTotalNeg = $mtdDsatComments.Count
$vocTotalPos = $mtdSatComments.Count

# ?? DSAT DEEP DIVE - TOP 3 ACTIONABLE ITEMS (current month) ????????????????
# Uses same source as VOC - only DSAT responses with comments
# Normalise L1 same as DSAT drivers
$mtdDsatNorm = $mtdDsatComments | ForEach-Object {
    $cat = $_.L1_CATEGORY
    if     ($cat -match 'Exam - System Based' -or $cat -match '^System Based') { $cat = 'System Based Assessment' }
    elseif ($cat -match 'Practice System')                                      { $cat = 'Practice Systems' }
    else                                                                        { $cat = 'Others' }
    [PSCustomObject]@{
        L1        = $cat
        L1_Orig   = $_.L1_CATEGORY
        Course    = ($_.EXT_COURSECODE -replace '_\d{4}$','')
        Comment   = if ($_.COMMENT1.Trim() -ne '') { $_.COMMENT1.Trim() } else { $_.COMMENT2.Trim() }
        Rating    = $_.RATING
    }
}

# Build top 3 deep dive items: pull counts directly from $negRows (same as VOC)
# Also enrich with L1 + course data from the original rows
$deepDiveSource = @()
$vocBucketOrder = @('Issue Not Resolved','Exam / System Issues','Agent Knowledge','Slow / Delayed Response','Case Closed Prematurely','Exam Result Dispute')

# Build a lookup from DISPLAYID -> normalised L1 + course for enrichment
$mtdDsatNorm = $mtdDsatComments | ForEach-Object {
    $cat = $_.L1_CATEGORY
    if     ($cat -match 'Exam - System Based' -or $cat -match '^System Based') { $cat = 'System Based Assessment' }
    elseif ($cat -match 'Practice System')                                      { $cat = 'Practice Systems' }
    else                                                                        { $cat = 'Others' }
    [PSCustomObject]@{
        DISPLAYID = $_.DISPLAYID
        L1        = $cat
        Course    = ($_.EXT_COURSECODE -replace '_\d{4}$','')
        Comment   = if ($_.COMMENT1.Trim() -ne '') { $_.COMMENT1.Trim() } else { $_.COMMENT2.Trim() }
    }
}
$normLookup = @{}
foreach ($r in $mtdDsatNorm) { $normLookup[$r.DISPLAYID] = $r }

foreach ($bucketKey in $vocBucketOrder) {
    $rows = $negRows[$bucketKey]
    if (-not $rows -or $rows.Count -eq 0) { continue }
    # Enrich with L1 + course
    $enriched = $rows | ForEach-Object {
        $n = $normLookup[$_.DISPLAYID]
        [PSCustomObject]@{
            L1     = if ($n) { $n.L1 } else { 'Others' }
            Course = if ($n) { $n.Course } else { '' }
        }
    }
    $topL1     = $enriched | Group-Object L1     | Sort-Object Count -Descending | Select-Object -First 1
    $topCourse = $enriched | Where-Object { $_.Course -notin @('','None','N/A') } |
                 Group-Object Course | Sort-Object Count -Descending | Select-Object -First 1
    $deepDiveSource += [PSCustomObject]@{
        Bucket     = $bucketKey
        L1         = $topL1.Name
        Count      = $rows.Count
        TopCourse  = if ($topCourse) { $topCourse.Name  } else { $null }
        TopCourseN = if ($topCourse) { $topCourse.Count } else { 0 }
        Color      = $negBuckets[$bucketKey].color
    }
    if ($deepDiveSource.Count -ge 3) { break }
}

# Root cause + action lookup per bucket
$rootCauses = @{
    'Issue Not Resolved'      = "Cases are being closed without a confirmed fix. Learners recontact with the same unresolved issue, indicating premature closure or insufficient follow-through by the agent."
    'Exam / System Issues'    = "Learners encounter technical failures during live certification attempts - timer misreads, unexpected logouts, and no clear rescue path - leaving them without a resolution at a critical moment."
    'Agent Knowledge'         = "Agents are providing generic or inaccurate responses that do not address the specific query. Gaps in product or process knowledge are visible to the customer."
    'Slow / Delayed Response' = "Response times for ILT and practice system issues are exceeding customer expectations. Delays of 10-15+ hours during time-sensitive delivery windows drive strong dissatisfaction."
    'Case Closed Prematurely' = "Tickets are being marked resolved before the customer confirms the issue is fixed. This forces learners to reopen cases and signals poor closure quality."
    'Exam Result Dispute'     = "Learners disputing exam outcomes receive no meaningful review or explanation. The lack of a visible escalation path amplifies frustration and damages trust."
}
$actions = @{
    'Issue Not Resolved'      = "Enforce a resolution confirmation step before closure - agent must log the fix applied and confirm with the learner. Flag any case reopened within 72 hours for supervisor review."
    'Exam / System Issues'    = "Create a dedicated fast-track queue for live exam technical issues with a 2-hour SLA. Equip agents with a standard rescue protocol (extend attempt, escalate to proctor team) to reduce resolution ambiguity."
    'Agent Knowledge'         = "Schedule targeted refresher sessions on top DSAT courses ($(if($deepDiveSource | Where-Object { $_.Bucket -eq 'Agent Knowledge' -and $_.TopCourse }){ ($deepDiveSource | Where-Object { $_.Bucket -eq 'Agent Knowledge' }).TopCourse }else{'SBA and PS topics'})). Introduce peer review of responses flagged for low CSAT before closure."
    'Slow / Delayed Response' = "Set a 4-hour IRT target for ILT and practice system cases. Introduce an auto-escalation trigger at 6 hours with supervisor notification to prevent SLA breaches going unnoticed."
    'Case Closed Prematurely' = "Add a mandatory learner confirmation field to the closure workflow. Implement a 24-hour reopen window with automatic DSAT risk flag for cases closed without learner sign-off."
    'Exam Result Dispute'     = "Publish a clear exam result review policy accessible at first contact. Route dispute cases to a dedicated senior agent with a defined 48-hour response commitment."
}

# Build deep dive HTML
$deepDiveHtml = ""
$rank = 1
foreach ($item in $deepDiveSource) {
    $rootCause = $rootCauses[$item.Bucket]
    $action    = $actions[$item.Bucket]
    $courseNote = if ($item.TopCourse -and $item.TopCourse -notin @('','None','N/A')) {
        "<span style='font-size:11px;background:#f1f5f9;color:#64748b;padding:2px 8px;border-radius:8px;margin-left:6px;'>Top course: <strong>$($item.TopCourse)</strong> ($($item.TopCourseN) DSATs)</span>"
    } else { "" }
    $pct = [math]::Round($item.Count / $mtdDsatComments.Count * 100)
    $id  = "dd$rank"
    $deepDiveHtml += @"
      <div style="background:#fff;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,0.07);overflow:hidden;margin-bottom:14px;border:1px solid #e2e8f0;">
        <div onclick="toggleDD('$id')" style="display:flex;align-items:center;gap:14px;padding:16px 20px;cursor:pointer;user-select:none;">
          <div style="width:32px;height:32px;border-radius:50%;background:$($item.Color);color:#fff;font-size:14px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;">#$rank</div>
          <div style="flex:1;">
            <div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px;">
              <span style="font-size:14px;font-weight:700;color:#1e293b;">$($item.Bucket)</span>
              <span style="font-size:11px;color:#94a3b8;">$($item.L1)</span>
              $courseNote
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:12px;flex-shrink:0;">
            <span style="font-size:13px;font-weight:700;color:$($item.Color);">$($item.Count) DSATs <span style="font-size:10px;color:#94a3b8;font-weight:400;">($pct%)</span></span>
            <span id="arr$rank" style="font-size:16px;color:#94a3b8;transition:transform 0.2s;">+</span>
          </div>
        </div>
        <div id="$id" style="display:none;padding:0 20px 18px;border-top:1px solid #f1f5f9;">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:14px;">
            <div style="background:#f8fafc;border-radius:8px;padding:14px 16px;">
              <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#94a3b8;margin-bottom:8px;">Root Cause</div>
              <div style="font-size:12px;color:#374151;line-height:1.7;">$rootCause</div>
            </div>
            <div style="background:#fffbeb;border-radius:8px;padding:14px 16px;border-left:3px solid #f59e0b;">
              <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#f59e0b;margin-bottom:8px;">Recommended Action</div>
              <div style="font-size:12px;color:#374151;line-height:1.7;">$action</div>
            </div>
          </div>
        </div>
      </div>
"@
    $rank++
}

# ?? VERBATIM COMMENTS (DSAT only, non-empty) ????????????????
$verbatim = @()
$dsatWithComments = $dsatAll | Where-Object { $_.COMMENT1.Trim() -ne '' -or $_.COMMENT2.Trim() -ne '' }
foreach ($r in ($dsatWithComments | Sort-Object { Get-Random } | Select-Object -First 6)) {
    $comment = if ($r.COMMENT1.Trim() -ne '') { $r.COMMENT1.Trim() } else { $r.COMMENT2.Trim() }
    if ($comment.Length -gt 220) { $comment = $comment.Substring(0,217) + '...' }
    $verbatim += [PSCustomObject]@{
        Rating   = $r.RATING
        Category = $r.L1_CATEGORY
        Course   = ($r.EXT_COURSECODE -replace '_\d{4}$','')
        Comment  = $comment
    }
}

# ?? AUTO INSIGHTS ????????????????????????????????????????????
$insightCSAT   = if ($ytdCSAT -ge $CSAT_TARGET) { "YTD CSAT at <strong>$ytdCSAT%</strong> - $csatVsTargetFmt." } else { "YTD CSAT at <strong>$ytdCSAT%</strong> - $csatVsTargetFmt. Needs attention." }
$insightMTD    = if ($mtdVsPrev -ge 0) { "$CurrentMonthShort MTD CSAT at <strong>$mtdCSAT%</strong> - $mtdVsPrevFmt." } else { "$CurrentMonthShort MTD CSAT at <strong>$mtdCSAT%</strong> - $mtdVsPrevFmt." }
$insightSplit  = "SBA CSAT at <strong>$sbaCSAT%</strong> vs Practice Systems at <strong>$psCSAT%</strong>. SBA accounts for $sbaShare% of all surveys."
$topDsatDriver = if ($dsatByL1) { $dsatByL1[0].Name } else { "N/A" }
$topDsatCount  = if ($dsatByL1) { $dsatByL1[0].Count } else { 0 }
$insightDSAT   = "Top DSAT driver: <strong>$topDsatDriver</strong> with $topDsatCount responses ($([math]::Round($topDsatCount/$ytdDSAT*100))% of all DSATs)."

# ?? BUILD HTML COMPONENTS ????????????????????????????????????

# Monthly trend bars (SVG)
$barCount   = $monthlyTrend.Count
$svgW       = 560
$barW       = 56
$barGap     = 14
$startX     = 30
$svgH       = 220
# y: 80% = y140, 100% = y0 => 1% = 3.5px, y = 140 - (val-80)*7 ... scale: 60-100 range, height=160
# y axis: 60% at y=160, 100% at y=0 => 1% = 4px
$trendBars = ""; $trendLabels = ""; $trendNLabels = ""; $trendMonthLabels = ""
$ci = 0
foreach ($m in $monthlyTrend) {
    $cx  = $startX + $ci * ($barW + $barGap) + $barW/2
    $bx  = $startX + $ci * ($barW + $barGap)
    $yTop = [math]::Round(160 - ($m.Pct - 60) * 4)
    $h    = 160 - $yTop
    if ($h -lt 2) { $h = 2; $yTop = 158 }
    $fill = if ($m.IsMtd) { '#007DB8' } elseif ($m.Pct -lt $CSAT_TARGET) { '#e05252' } else { '#38b6e8' }
    $labelFill = $fill
    $trendBars         += "    <rect x=""$bx"" y=""$yTop"" width=""$barW"" height=""$h"" rx=""4"" fill=""$fill"" opacity=""$(if($m.IsMtd){'1.0'}else{'0.85'})""/>`n"
    $trendLabels       += "    <text x=""$cx"" y=""$($yTop-5)"" text-anchor=""middle"" font-size=""11"" font-weight=""bold"" fill=""$labelFill"">$($m.Pct)%</text>`n"
    $trendNLabels      += "    <text x=""$cx"" y=""175"" text-anchor=""middle"" font-size=""9"" fill=""#94a3b8"">n=$($m.N)</text>`n"
    $moColor = if ($m.IsMtd) { '#007DB8' } else { '#64748b' }
    $trendMonthLabels  += "    <text x=""$cx"" y=""192"" text-anchor=""middle"" font-size=""11"" font-weight=""600"" fill=""$moColor"">$(if($m.IsMtd){"$($m.Short) MTD"}else{$m.Short})</text>`n"
    $ci++
}
$targetY = [math]::Round(160 - ($CSAT_TARGET - 60) * 4)

# DSAT driver bars
$dsatDriverHtml = ""
foreach ($d in $dsatByL1) {
    $pct = [math]::Round($d.Count / $ytdDSAT * 100)
    $w   = [math]::Round($d.Count / $maxDsatL1 * 100, 1)
    $dsatDriverHtml += @"
        <div style="margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;">
            <span style="font-size:13px;font-weight:600;color:#1e293b;">$($d.Name)</span>
            <span style="font-size:12px;font-weight:700;color:#e05252;">$($d.Count) <span style="font-size:10px;color:#94a3b8;font-weight:400;">($pct%)</span></span>
          </div>
          <div style="background:#e2e8f0;border-radius:4px;height:10px;overflow:hidden;">
            <div style="width:$($w)%;height:100%;background:#e05252;border-radius:4px;opacity:0.8;"></div>
          </div>
        </div>
"@
}

# Course DSAT bars
$courseColors = @('#e05252','#f97316','#f59e0b','#94a3b8','#94a3b8')
$courseDsatHtml = ""
$ci2 = 0
foreach ($c in $dsatCourses) {
    $pct = [math]::Round($c.Count / $ytdDSAT * 100)
    $w   = [math]::Round($c.Count / $maxDsatCourse * 100, 1)
    $col = $courseColors[$ci2]
    $courseDsatHtml += @"
        <div style="margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;">
            <span style="font-size:13px;font-weight:700;color:#1e293b;">$($c.Name)</span>
            <span style="font-size:12px;font-weight:700;color:$col;">$($c.Count) DSATs <span style="font-size:10px;color:#94a3b8;font-weight:400;">($pct% of total)</span></span>
          </div>
          <div style="background:#e2e8f0;border-radius:4px;height:10px;overflow:hidden;">
            <div style="width:$($w)%;height:100%;background:$col;border-radius:4px;opacity:0.85;"></div>
          </div>
        </div>
"@
    $ci2++
}

# Verbatim cards
$verbatimHtml = ""
foreach ($v in $verbatim) {
    $borderColor = if ($v.Rating -eq 'Very Dissatisfied') { '#e05252' } else { '#f97316' }
    $badge = if ($v.Rating -eq 'Very Dissatisfied') { 'background:#fef2f2;color:#e05252;' } else { 'background:#fff7ed;color:#f97316;' }
    $courseTag = if ($v.Course -notin @('','None','N/A')) { "<span style='font-size:10px;background:#f1f5f9;color:#64748b;padding:2px 7px;border-radius:8px;'>$($v.Course)</span>" } else { "" }
    $verbatimHtml += @"
        <div style="background:#fff;border-radius:10px;padding:16px 18px;border-left:4px solid $borderColor;box-shadow:0 1px 4px rgba(0,0,0,0.06);margin-bottom:14px;">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
            <span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:8px;$badge">$($v.Rating)</span>
            <span style="font-size:10px;color:#94a3b8;">$($v.Category)</span>
            $courseTag
          </div>
          <div style="font-size:13px;color:#374151;line-height:1.6;font-style:italic;">"$($v.Comment)"</div>
        </div>
"@
}

# SBA vs PS split bar
$sbaBarW = [math]::Round($sbaCSAT / 100 * 100, 1)
$psBarW  = [math]::Round($psCSAT  / 100 * 100, 1)
$iltBarW = [math]::Round($iltCSAT / 100 * 100, 1)
$sbaFill = if ($sbaCSAT -ge $CSAT_TARGET) { '#007DB8' } else { '#e05252' }
$psFill  = if ($psCSAT  -ge $CSAT_TARGET) { '#38b6e8' } else { '#f97316' }
$iltFill = if ($iltCSAT -ge $CSAT_TARGET) { '#10b981' } else { '#f97316' }

# ?? WRITE HTML ????????????????????????????????????????????????
$outputFile = "$OutputDir\GTS_CSAT_Bulletin.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>GTS ESM - CSAT Bulletin $CurrentMonthShort 2026</title>
  <style>
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:'Segoe UI',Arial,sans-serif; background:#f1f5f9; color:#1e293b; padding:36px 32px; max-width:1100px; margin:0 auto; }
    .page-header { margin-bottom:32px; padding-bottom:18px; border-bottom:3px solid #007DB8; display:flex; justify-content:space-between; align-items:flex-end; }
    .page-header h1 { font-size:24px; font-weight:700; color:#007DB8; }
    .page-header .subtitle { font-size:13px; color:#64748b; margin-top:4px; }
    .page-header .datestamp { font-size:12px; color:#94a3b8; text-align:right; line-height:1.7; }
    .kpi-row { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin-bottom:32px; }
    .kpi-card { background:#fff; border-radius:12px; padding:20px 18px 16px; box-shadow:0 2px 8px rgba(0,0,0,0.06); border-top:4px solid var(--accent); }
    .kpi-card.blue  { --accent:#007DB8; } .kpi-card.lblue { --accent:#38b6e8; }
    .kpi-card.red   { --accent:#e05252; } .kpi-card.amber { --accent:#f97316; }
    .kpi-label { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#94a3b8; margin-bottom:8px; }
    .kpi-value { font-size:34px; font-weight:700; color:var(--accent); line-height:1; }
    .kpi-note  { font-size:11px; color:#94a3b8; margin-top:8px; }
    .kpi-note strong { color:#64748b; }
    .insights { background:#fff; border-radius:12px; padding:18px 22px; margin-bottom:32px; border-left:4px solid #007DB8; box-shadow:0 2px 8px rgba(0,0,0,0.06); }
    .insights-title { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#007DB8; margin-bottom:12px; }
    .insight-item { display:flex; align-items:flex-start; gap:10px; margin-bottom:8px; font-size:12px; color:#374151; line-height:1.6; }
    .insight-item:last-child { margin-bottom:0; }
    .i-dot { width:7px; height:7px; border-radius:50%; flex-shrink:0; margin-top:5px; }
    .i-dot.blue { background:#007DB8; } .i-dot.green { background:#22c55e; } .i-dot.red { background:#e05252; } .i-dot.amber { background:#f97316; }
    .section { margin-bottom:32px; }
    .section-header { display:flex; align-items:center; gap:12px; margin-bottom:18px; }
    .section-header h2 { font-size:13px; font-weight:700; text-transform:uppercase; letter-spacing:1.2px; color:#007DB8; white-space:nowrap; }
    .section-rule { flex:1; height:1px; background:linear-gradient(to right,#007DB8,transparent); }
    .card { background:#fff; border-radius:12px; padding:24px; box-shadow:0 2px 8px rgba(0,0,0,0.06); }
    .two-col { display:grid; grid-template-columns:1fr 1fr; gap:20px; }
    .three-col { display:grid; grid-template-columns:1fr 1fr 1fr; gap:20px; }
    .split-row { display:flex; align-items:center; gap:14px; margin-bottom:18px; }
    .split-label { width:130px; font-size:13px; font-weight:600; color:#374151; flex-shrink:0; }
    .split-track { flex:1; background:#e2e8f0; border-radius:6px; height:22px; overflow:hidden; position:relative; }
    .split-fill { height:100%; border-radius:6px; display:flex; align-items:center; padding-left:10px; transition:width 0.4s; }
    .split-fill span { font-size:11px; font-weight:700; color:#fff; }
    .split-score { width:52px; text-align:right; font-size:15px; font-weight:700; flex-shrink:0; }
    .split-meta { font-size:11px; color:#94a3b8; margin-top:2px; }
    .tab-bar { display:flex; gap:8px; margin-bottom:16px; }
    .tab { padding:5px 14px; border-radius:20px; font-size:12px; font-weight:600; cursor:pointer; border:none; background:#e2e8f0; color:#64748b; transition:all 0.2s; }
    .tab.active { background:#007DB8; color:#fff; }
    .page-footer { margin-top:16px; font-size:11px; color:#cbd5e1; text-align:center; padding-top:14px; border-top:1px solid #e2e8f0; }
  </style>
  <script>
    function toggleDD(id) {
      var el  = document.getElementById(id);
      var num = id.replace('dd','');
      var arr = document.getElementById('arr' + num);
      if (el.style.display === 'none') {
        el.style.display = 'block';
        arr.textContent  = '-';
        arr.style.color  = '#007DB8';
      } else {
        el.style.display = 'none';
        arr.textContent  = '+';
        arr.style.color  = '#94a3b8';
      }
    }
  </script>
</head>
<body>

  <div class="page-header">
    <div>
      <h1>GTS ESM - CSAT Bulletin</h1>
      <div class="subtitle">Global Training Support &nbsp;&middot;&nbsp; Customer Satisfaction Report &nbsp;&middot;&nbsp; YTD 2026</div>
    </div>
    <div class="datestamp">As of $ReportDate<br>Source: SERVICE_CLOUD &nbsp;&middot;&nbsp; $($ytdN.ToString('N0')) surveys YTD</div>
  </div>

  <!-- KPI STRIP -->
  <div class="kpi-row">
    <div class="kpi-card blue">
      <div class="kpi-label">YTD CSAT</div>
      <div class="kpi-value">$ytdCSAT%</div>
      <div class="kpi-note"><strong>$($ytdN.ToString('N0')) surveys</strong> &nbsp;&middot;&nbsp; $csatVsTargetFmt</div>
    </div>
    <div class="kpi-card lblue">
      <div class="kpi-label">$CurrentMonthShort MTD CSAT</div>
      <div class="kpi-value">$mtdCSAT%</div>
      <div class="kpi-note"><strong>$($mtdN.ToString('N0')) surveys</strong> &nbsp;&middot;&nbsp; $mtdVsPrevFmt</div>
    </div>
    <div class="kpi-card amber">
      <div class="kpi-label">YTD DSATs</div>
      <div class="kpi-value">$($ytdDSAT.ToString('N0'))</div>
      <div class="kpi-note"><strong>$([math]::Round($ytdDSAT/$ytdN*100,1))%</strong> of all surveys</div>
    </div>
    <div class="kpi-card red">
      <div class="kpi-label">$CurrentMonthShort DSATs</div>
      <div class="kpi-value">$($mtdDSAT.ToString('N0'))</div>
      <div class="kpi-note"><strong>$([math]::Round($mtdDSAT/$mtdN*100,1))%</strong> of $CurrentMonthShort surveys</div>
    </div>
  </div>

  <!-- INSIGHTS -->
  <div class="insights">
    <div class="insights-title">Key Insights</div>
    <div class="insight-item"><div class="i-dot $(if($ytdCSAT -ge $CSAT_TARGET){'green'}else{'red'})"></div><div>$insightCSAT</div></div>
    <div class="insight-item"><div class="i-dot $(if($mtdVsPrev -ge 0){'green'}else{'amber'})"></div><div>$insightMTD</div></div>
    <div class="insight-item"><div class="i-dot blue"></div><div>$insightSplit</div></div>
    <div class="insight-item"><div class="i-dot red"></div><div>$insightDSAT</div></div>
  </div>

  <!-- MONTHLY TREND -->
  <div class="section">
    <div class="section-header"><h2>Monthly CSAT Trend</h2><div class="section-rule"></div></div>
    <div class="card">
      <svg viewBox="0 0 $(($barCount * ($barW + $barGap)) + $startX * 2) 210" style="width:100%;height:210px;" preserveAspectRatio="none">
        <line x1="$startX" y1="0"   x2="$(($barCount * ($barW + $barGap)) + $startX)" y2="0"   stroke="#f1f5f9" stroke-width="1"/>
        <line x1="$startX" y1="40"  x2="$(($barCount * ($barW + $barGap)) + $startX)" y2="40"  stroke="#f1f5f9" stroke-width="1"/>
        <line x1="$startX" y1="80"  x2="$(($barCount * ($barW + $barGap)) + $startX)" y2="80"  stroke="#f1f5f9" stroke-width="1"/>
        <line x1="$startX" y1="120" x2="$(($barCount * ($barW + $barGap)) + $startX)" y2="120" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="$startX" y1="160" x2="$(($barCount * ($barW + $barGap)) + $startX)" y2="160" stroke="#f1f5f9" stroke-width="1"/>
        <text x="$(($startX - 4))" y="4"   text-anchor="end" font-size="9" fill="#94a3b8">100%</text>
        <text x="$(($startX - 4))" y="44"  text-anchor="end" font-size="9" fill="#94a3b8">90%</text>
        <text x="$(($startX - 4))" y="84"  text-anchor="end" font-size="9" fill="#22c55e" font-weight="bold">80%</text>
        <text x="$(($startX - 4))" y="124" text-anchor="end" font-size="9" fill="#94a3b8">70%</text>
        <text x="$(($startX - 4))" y="164" text-anchor="end" font-size="9" fill="#94a3b8">60%</text>
        <line x1="$startX" y1="$targetY" x2="$(($barCount * ($barW + $barGap)) + $startX)" y2="$targetY" stroke="#22c55e" stroke-width="1.5" stroke-dasharray="6,3"/>
        <text x="$(($barCount * ($barW + $barGap)) + $startX + 4)" y="$($targetY + 4)" font-size="9" fill="#22c55e" font-weight="bold">Target</text>
        $trendBars
        $trendLabels
        $trendNLabels
        $trendMonthLabels
      </svg>
      <div style="display:flex;gap:20px;margin-top:10px;padding-top:10px;border-top:1px solid #f1f5f9;flex-wrap:wrap;">
        <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:#64748b;"><div style="width:14px;height:14px;background:#38b6e8;border-radius:3px;opacity:0.85;"></div> CSAT Monthly</div>
        <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:#64748b;"><div style="width:14px;height:14px;background:#007DB8;border-radius:3px;"></div> $CurrentMonthShort MTD</div>
        <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:#64748b;"><div style="width:14px;height:14px;background:#e05252;border-radius:3px;opacity:0.8;"></div> Below target</div>
        <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:#64748b;"><div style="width:24px;height:1px;border-top:2px dashed #22c55e;"></div> Target ($CSAT_TARGET%)</div>
      </div>
    </div>
  </div>

  <!-- CSAT SPLIT -->
  <div class="section">
    <div class="section-header"><h2>CSAT Split - SBA vs Practice Systems vs ILT</h2><div class="section-rule"></div></div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;">
      <!-- SBA -->
      <div class="card" style="border-top:4px solid $sbaFill;">
        <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:10px;">
          <span style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:#64748b;">System Based (SBA)</span>
          <span style="font-size:11px;color:#94a3b8;">$sbaShare% of all surveys</span>
        </div>
        <div style="font-size:36px;font-weight:700;color:$sbaFill;line-height:1;margin-bottom:10px;">$sbaCSAT%</div>
        <div style="background:#e2e8f0;border-radius:6px;height:10px;overflow:hidden;margin-bottom:10px;">
          <div style="width:$($sbaBarW)%;height:100%;background:$sbaFill;border-radius:6px;"></div>
        </div>
        <div style="display:flex;gap:16px;font-size:11px;color:#64748b;">
          <span><strong style="color:#1e293b;">$($sbaN.ToString('N0'))</strong> surveys</span>
          <span><strong style="color:#e05252;">$([math]::Round($sbaDSAT/$sbaN*100,1))%</strong> DSAT rate</span>
        </div>
      </div>
      <!-- Practice Systems -->
      <div class="card" style="border-top:4px solid $psFill;">
        <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:10px;">
          <span style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:#64748b;">Practice Systems</span>
          <span style="font-size:11px;color:#94a3b8;">$psShare% of all surveys</span>
        </div>
        <div style="font-size:36px;font-weight:700;color:$psFill;line-height:1;margin-bottom:10px;">$psCSAT%</div>
        <div style="background:#e2e8f0;border-radius:6px;height:10px;overflow:hidden;margin-bottom:10px;">
          <div style="width:$($psBarW)%;height:100%;background:$psFill;border-radius:6px;"></div>
        </div>
        <div style="display:flex;gap:16px;font-size:11px;color:#64748b;">
          <span><strong style="color:#1e293b;">$($psN.ToString('N0'))</strong> surveys</span>
          <span><strong style="color:#e05252;">$([math]::Round($psDSAT/$psN*100,1))%</strong> DSAT rate</span>
        </div>
      </div>
      <!-- ILT -->
      <div class="card" style="border-top:4px solid $iltFill;">
        <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:10px;">
          <span style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:#64748b;">ILT</span>
          <span style="font-size:11px;color:#94a3b8;">$iltShare% of all surveys</span>
        </div>
        <div style="font-size:36px;font-weight:700;color:$iltFill;line-height:1;margin-bottom:10px;">$iltCSAT%</div>
        <div style="background:#e2e8f0;border-radius:6px;height:10px;overflow:hidden;margin-bottom:10px;">
          <div style="width:$($iltBarW)%;height:100%;background:$iltFill;border-radius:6px;"></div>
        </div>
        <div style="display:flex;gap:16px;font-size:11px;color:#64748b;">
          <span><strong style="color:#1e293b;">$($iltN.ToString('N0'))</strong> surveys</span>
          <span><strong style="color:#e05252;">$([math]::Round($iltDSAT/$iltN*100,1))%</strong> DSAT rate</span>
        </div>
      </div>
    </div>
  </div>

  <!-- DSAT SECTION -->
  <div class="section">
    <div class="section-header"><h2>DSAT Analysis</h2><div class="section-rule"></div></div>
    <div class="two-col">
      <div class="card">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#94a3b8;margin-bottom:16px;">By Category</div>
        $dsatDriverHtml
      </div>
      <div class="card">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#94a3b8;margin-bottom:16px;">Top Courses by DSAT Volume</div>
        $courseDsatHtml
      </div>
    </div>
  </div>

  <!-- VOICE OF CUSTOMER -->
  <div class="section">
    <div class="section-header"><h2>Voice of Customer - $CurrentMonthShort Comment Themes</h2><div class="section-rule"></div></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
      <!-- POSITIVES -->
      <div>
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
          <div style="width:10px;height:10px;border-radius:50%;background:#22c55e;flex-shrink:0;"></div>
          <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#22c55e;">What Customers Liked</span>
          <span style="font-size:10px;color:#94a3b8;margin-left:4px;">$vocTotalPos SAT comments</span>
        </div>
        $posColHtml
      </div>
      <!-- NEGATIVES -->
      <div>
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
          <div style="width:10px;height:10px;border-radius:50%;background:#e05252;flex-shrink:0;"></div>
          <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#e05252;">What Customers Complained About</span>
          <span style="font-size:10px;color:#94a3b8;margin-left:4px;">$vocTotalNeg DSAT comments</span>
        </div>
        $negColHtml
      </div>
    </div>
  </div>

  <!-- DSAT DEEP DIVE -->
  <div class="section">
    <div class="section-header"><h2>DSAT Deep Dive - Top 3 Action Items</h2><div class="section-rule"></div></div>
    <div style="font-size:12px;color:#64748b;margin-bottom:14px;">Based on $CurrentMonthShort DSAT responses. Click any item to expand root cause and recommended action.</div>
    $deepDiveHtml
  </div>

  <!-- VERBATIM COMMENTS -->
  <div class="section">
    <div class="section-header"><h2>Verbatim - DSAT Responses</h2><div class="section-rule"></div></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
      $verbatimHtml
    </div>
  </div>

  <div class="page-footer">
    Generated $ReportDate &nbsp;&middot;&nbsp; GTS ESM Operations &nbsp;&middot;&nbsp; Surveys: $(Split-Path $SurveysXlsx -Leaf) &nbsp;&middot;&nbsp; Cases: $(Split-Path $CasesXlsx -Leaf)
  </div>

</body>
</html>
"@

[System.IO.File]::WriteAllText($outputFile, $html, [System.Text.UTF8Encoding]::new($false))
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
