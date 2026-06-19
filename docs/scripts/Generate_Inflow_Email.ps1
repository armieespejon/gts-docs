# ============================================================
# GTS ESM Inflow & Closure Email Generator
# Auto-opens Outlook with summary email ready to send
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_Inflow_Email.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$To      = "recipient1@sap.com; recipient2@sap.com"   # separate with semicolons
$CC      = ""                                           # optional
$ReportPath = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\GTS_Daily_Inflow_June.html"

$CasesXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx" | Where-Object { $_.Name -notmatch 'with closed' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
# -- END CONFIG -----------------------------------------------

# Derive report date
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

$W1Start = 1;  $W1End = 7
$W2Start = 8;  $W2End = 14
$W3Start = 15

# Export xlsx to CSV
$CasesCsv = "$env:TEMP\gts_cases_email.csv"
Write-Host "Exporting data..." -ForegroundColor Yellow
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($CasesXlsx)
$wb.SaveAs($CasesCsv, 6)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

$peekContent = Get-Content $CasesCsv -TotalCount 3
$HasClosures = ($peekContent -join ' ') -match 'COMPLETED_ON|CLOSED_ON'

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

$deKeywords = @{
    'mjay'='Mritunjay Kumar Singh'; 'mritunjay'='Mritunjay Kumar Singh'
    'shub'='Shubham Thakur';        'shubham'='Shubham Thakur'
    'ani'='Anjali Tripathi';        'anjali'='Anjali Tripathi'
    'swarna'='Swarna Gupta';        'kriti'='Kriti Bhalla'
    'naren'='Narendran Raja';       'narendran'='Narendran Raja'
}
function Resolve-Processor($row) {
    $procId = $row.PROCESSOR_ID; $misc = $row.EXT_MISCINFO.Trim().ToLower(); $name = $row.PROCESSOR_NAME.Trim()
    if ($procId -eq 'SAP_GERMANY_USER') {
        foreach ($key in $deKeywords.Keys) { if ($misc -match $key) { return $deKeywords[$key] } }
        return 'Germany (Unassigned)'
    }
    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq 'None') { return 'Unassigned' }
    return $name
}

# PH team only for email summary
$phTeam = @('Quinto','Paulino','Valdez','Nobela','Sinag','Mendoza','Buenafe','Cada','Santos','Chua','Tulagan','Evidente','Cruz')

$monthCases = $cases | Where-Object { $_.MONTH -eq $Month }

# Build per-agent totals
$agentData = @{}
foreach ($row in $monthCases) {
    $p = Resolve-Processor $row
    $isph = $false
    foreach ($ln in $phTeam) { if ($p -match $ln) { $isph = $true; break } }
    if (-not $isph) { continue }
    if (-not $agentData.ContainsKey($p)) { $agentData[$p] = @{ Inflow=0; Close=0 } }
    $agentData[$p].Inflow++
}

if ($HasClosures) {
    foreach ($row in $cases) {
        $doneDate = if ($row.COMPLETED_ON -ne '') { $row.COMPLETED_ON } else { $row.CLOSED_ON }
        if ([string]::IsNullOrWhiteSpace($doneDate)) { continue }
        try { $dt = [datetime]::Parse($doneDate) } catch { continue }
        if ($dt.ToString('MMMM') -ne $Month) { continue }
        $p = Resolve-Processor $row
        $isph = $false
        foreach ($ln in $phTeam) { if ($p -match $ln) { $isph = $true; break } }
        if (-not $isph) { continue }
        if (-not $agentData.ContainsKey($p)) { $agentData[$p] = @{ Inflow=0; Close=0 } }
        $agentData[$p].Close++
    }
}

# Team totals
$teamInflow = 0; $teamClose = 0
foreach ($v in $agentData.Values) { $teamInflow += $v.Inflow; $teamClose += $v.Close }
$teamRate   = if ($teamInflow -gt 0) { [math]::Round($teamClose / $teamInflow * 100) } else { 0 }
$activeCount = ($cases | Where-Object { $_.STATUS -notin @('Closed','Completed') }).Count

# Top performer (highest closure rate, min 10 inflow)
$topAgent = $agentData.GetEnumerator() | Where-Object { $_.Value.Inflow -ge 10 } | ForEach-Object {
    $rate = if ($_.Value.Inflow -gt 0) { [math]::Round($_.Value.Close / $_.Value.Inflow * 100) } else { 0 }
    [PSCustomObject]@{ Name=$_.Key; Inflow=$_.Value.Inflow; Close=$_.Value.Close; Rate=$rate }
} | Sort-Object Rate -Descending | Select-Object -First 1

# Build agent table rows
$tableRows = ""
$sorted = $agentData.GetEnumerator() | ForEach-Object {
    $rate = if ($_.Value.Inflow -gt 0) { [math]::Round($_.Value.Close / $_.Value.Inflow * 100) } else { 0 }
    [PSCustomObject]@{ Name=$_.Key; Inflow=$_.Value.Inflow; Close=$_.Value.Close; Rate=$rate }
} | Sort-Object Inflow -Descending

foreach ($a in $sorted) {
    $rateColor = if ($a.Rate -ge 100) { '#166534' } elseif ($a.Rate -ge 80) { '#854d0e' } else { '#991b1b' }
    $rateBg    = if ($a.Rate -ge 100) { '#dcfce7' } elseif ($a.Rate -ge 80) { '#fef9c3' } else { '#fee2e2' }
    $tableRows += "<tr style='border-bottom:1px solid #e5e7eb'>
      <td style='padding:7px 12px;font-weight:600;color:#1e3a5f'>$($a.Name)</td>
      <td style='padding:7px 12px;text-align:center'>$($a.Inflow)</td>
      <td style='padding:7px 12px;text-align:center'>$($a.Close)</td>
      <td style='padding:7px 12px;text-align:center;font-weight:700;color:$rateColor;background:$rateBg;border-radius:4px'>$($a.Rate)%</td>
    </tr>"
}

$rateColor = if ($teamRate -ge 100) { '#166534' } elseif ($teamRate -ge 80) { '#854d0e' } else { '#991b1b' }
$topLine   = if ($topAgent) { "Top performer this period: <strong>$($topAgent.Name)</strong> with a closure rate of <strong>$($topAgent.Rate)%</strong> ($($topAgent.Close) closed / $($topAgent.Inflow) inflow)." } else { "" }

$subject = "GTS ESM PH Team | Inflow & Closure Update | $ReportDate"

$body = @"
<html><body style="font-family:'Segoe UI',Arial,sans-serif;font-size:13px;color:#1a1a2e;max-width:700px">

<div style="background:#1e3a5f;padding:16px 20px;border-radius:6px 6px 0 0">
  <div style="font-size:18px;font-weight:700;color:#fff">GTS ESM PH Team</div>
  <div style="font-size:13px;color:#93c5fd;margin-top:4px">Inflow &amp; Closure Update &mdash; $ReportDate</div>
</div>

<div style="background:#f8fafc;padding:14px 20px;border:1px solid #e5e7eb;border-top:none">
  <p style="margin:0 0 10px">Hi team,</p>
  <p style="margin:0 0 14px">Here is the June inflow and closure summary as of <strong>$ReportDate</strong>. $topLine</p>

  <!-- KPI STRIP -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px">
    <tr>
      <td style="width:33%;padding:10px;background:#eff6ff;border-radius:6px;text-align:center;border:1px solid #bfdbfe">
        <div style="font-size:22px;font-weight:700;color:#1d4ed8">$($teamInflow.ToString('N0'))</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">Total Inflow</div>
      </td>
      <td style="width:4%"></td>
      <td style="width:33%;padding:10px;background:#f0fdf4;border-radius:6px;text-align:center;border:1px solid #bbf7d0">
        <div style="font-size:22px;font-weight:700;color:#166534">$($teamClose.ToString('N0'))</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">Total Closures</div>
      </td>
      <td style="width:4%"></td>
      <td style="width:26%;padding:10px;background:#fafafa;border-radius:6px;text-align:center;border:1px solid #e5e7eb">
        <div style="font-size:22px;font-weight:700;color:$rateColor">$teamRate%</div>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">Team Closure Rate</div>
      </td>
    </tr>
  </table>

  <!-- AGENT TABLE -->
  <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;font-size:12px">
    <thead>
      <tr style="background:#1e3a5f">
        <th style="padding:8px 12px;text-align:left;color:#fff;font-weight:600">Agent</th>
        <th style="padding:8px 12px;text-align:center;color:#93c5fd;font-weight:600">Inflow</th>
        <th style="padding:8px 12px;text-align:center;color:#6ee7b7;font-weight:600">Closures</th>
        <th style="padding:8px 12px;text-align:center;color:#fde68a;font-weight:600">Rate %</th>
      </tr>
    </thead>
    <tbody>
      $tableRows
    </tbody>
    <tfoot>
      <tr style="background:#f1f5f9;font-weight:700;border-top:2px solid #e5e7eb">
        <td style="padding:8px 12px;color:#1e3a5f">TEAM TOTAL</td>
        <td style="padding:8px 12px;text-align:center;color:#1d4ed8">$teamInflow</td>
        <td style="padding:8px 12px;text-align:center;color:#166534">$teamClose</td>
        <td style="padding:8px 12px;text-align:center;font-weight:700;color:$rateColor">$teamRate%</td>
      </tr>
    </tfoot>
  </table>

  <p style="margin:14px 0 0;font-size:11px;color:#6b7280">Full interactive report attached. Keep up the great work!</p>
</div>

<div style="background:#f1f5f9;padding:10px 20px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;font-size:11px;color:#9ca3af">
  GTS ESM Operations &mdash; Auto-generated report &mdash; $ReportDate
</div>

</body></html>
"@

Write-Host "Generating email file..." -ForegroundColor Yellow

# Save as .eml file — double-click to open in Outlook
$emlPath = "$env:TEMP\GTS_Inflow_Update_$($parsedDate.ToString('MMMdd')).eml"

$boundary = "----=_GTS_$(Get-Random)"

# Encode HTML body as base64
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
$bodyB64   = [Convert]::ToBase64String($bodyBytes)

# Read attachment as base64
$attachB64 = ""
$attachName = ""
if (Test-Path $ReportPath) {
    $attachBytes = [System.IO.File]::ReadAllBytes($ReportPath)
    $attachB64   = [Convert]::ToBase64String($attachBytes)
    $attachName  = Split-Path $ReportPath -Leaf
}

$toField = $To
$ccField = if ($CC) { "CC: $CC`r`n" } else { "" }

$eml = "From: GTS ESM Operations`r`n"
$eml += "To: $toField`r`n"
$eml += $ccField
$eml += "Subject: $subject`r`n"
$eml += "MIME-Version: 1.0`r`n"
$eml += "Content-Type: multipart/mixed; boundary=`"$boundary`"`r`n"
$eml += "`r`n"
$eml += "--$boundary`r`n"
$eml += "Content-Type: text/html; charset=UTF-8`r`n"
$eml += "Content-Transfer-Encoding: base64`r`n"
$eml += "`r`n"
$eml += $bodyB64
$eml += "`r`n"

if ($attachB64) {
    $eml += "--$boundary`r`n"
    $eml += "Content-Type: text/html; name=`"$attachName`"`r`n"
    $eml += "Content-Transfer-Encoding: base64`r`n"
    $eml += "Content-Disposition: attachment; filename=`"$attachName`"`r`n"
    $eml += "`r`n"
    $eml += $attachB64
    $eml += "`r`n"
}
$eml += "--$boundary--`r`n"

[System.IO.File]::WriteAllText($emlPath, $eml, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "Done! Email file saved. Opening now..." -ForegroundColor Green
Write-Host "  $emlPath" -ForegroundColor Yellow
Start-Process $emlPath
