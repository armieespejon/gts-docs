# ============================================================
# GTS ESM Agent Backlog Snapshot - Auto-Generator
#
# Usage:
#   1. Drop updated xlsx in Downloads
#   2. Run: powershell -ExecutionPolicy Bypass -File "Generate_Agent_Backlog.ps1"
# ============================================================

# AUTO-DETECT
$CasesXlsx = (Get-ChildItem "C:\Users\I535893\Downloads\GTS ESM Cases_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$OutputDir = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"

# Derive report date from filename
$baseName   = [System.IO.Path]::GetFileNameWithoutExtension($CasesXlsx)
$datePart   = $baseName -replace '^GTS ESM Cases_', ''
$parsedDate = $null
foreach ($fmt in @('MMMMd','MMMMdd','MMMd','MMMdd')) {
    try { $parsedDate = [datetime]::ParseExact($datePart, $fmt, [System.Globalization.CultureInfo]::InvariantCulture); break } catch {}
}
if (-not $parsedDate) { $parsedDate = [datetime]::Today }
$parsedDate  = $parsedDate.AddYears([datetime]::Today.Year - $parsedDate.Year)
$ReportDate  = $parsedDate.ToString("MMMM d, yyyy")
$today       = [datetime]::Today

$CasesCsv = "$env:TEMP\gts_cases_backlog.csv"

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

Write-Host "Computing agent backlog..." -ForegroundColor Yellow

# DE name resolution
$deNameMap = @{
    'mjay'   = 'Mritunjay Kumar Singh'
    'shub'   = 'Shubham Thakur'
    'ani'    = 'Anjali Tripathi'
    'swarna' = 'Swarna Gupta'
    'kriti'  = 'Kriti Bhalla'
    'naren'  = 'Narendran Raja'
}

$closedStatuses = @('Closed','Completed')
$activeStatuses = @('In Process','Customer Action','Awaiting Third Party Action','Open')

# Agent roster
$phAgents = @('Marviely Quinto','Justin Jan Paulino','Lemuel Von Valdez','Monica Frances Nobela',
              'Rendel Edison Sinag','Edrian Mendoza','Shaun Paul Buenafe','Saldy Cada',
              'Ana Lizelle Santos','Sam Tomi Chua','Nicole Tulagan','Marie Cristine Evidente','Ann Cruz')

# Map data names (Lastname, Firstname) resolved -> official display name
$phNameMap = @{
    'Justin Paulino' = 'Justin Jan Paulino'
    'Paul Buenafe'   = 'Shaun Paul Buenafe'
    'Sam Chua'       = 'Sam Tomi Chua'
}
$deAgents = @('Mritunjay Kumar Singh','Shubham Thakur','Anjali Tripathi','Swarna Gupta',
              'Kriti Bhalla','Narendran Raja')

# Resolve processor name
function Resolve-Name($row) {
    if ($row.PROCESSOR_ID -ne 'SAP_GERMANY_USER') {
        # PH: "Lastname, Firstname" -> "Firstname Lastname"
        $parts = $row.PROCESSOR_NAME -split ',\s*'
        $name  = if ($parts.Count -eq 2) { "$($parts[1].Trim()) $($parts[0].Trim())" } else { $row.PROCESSOR_NAME }
        # Apply official name map for truncated data names
        if ($phNameMap.ContainsKey($name)) { return $phNameMap[$name] }
        return $name
    }
    $misc = $row.EXT_MISCINFO.Trim().ToLower()
    foreach ($key in $deNameMap.Keys) {
        if ($misc -match "^$key") { return $deNameMap[$key] }
    }
    return $null
}

# Build per-agent stats
$agentStats = @{}
foreach ($row in $cases) {
    $name = Resolve-Name $row
    if (-not $name) { continue }
    $isRoster = ($phAgents + $deAgents) -contains $name
    if (-not $isRoster) { continue }
    if (-not $agentStats.ContainsKey($name)) {
        $agentStats[$name] = @{ Total=0; Active=0; Closed=0; Age0=0; Age16=0; Age30=0 }
    }
    $agentStats[$name].Total++
    if ($row.STATUS -in $closedStatuses) {
        $agentStats[$name].Closed++
    } else {
        $agentStats[$name].Active++
        $age = ($today - [datetime]::Parse($row.CREATED_ON_DATE)).Days
        if     ($age -le 15) { $agentStats[$name].Age0++ }
        elseif ($age -le 30) { $agentStats[$name].Age16++ }
        else                  { $agentStats[$name].Age30++ }
    }
}

# Team-level totals
$teamTotal  = 0; $teamActive = 0; $teamClosed = 0; $teamAge30 = 0
foreach ($k in $agentStats.Keys) {
    $teamTotal  += $agentStats[$k].Total
    $teamActive += $agentStats[$k].Active
    $teamClosed += $agentStats[$k].Closed
    $teamAge30  += $agentStats[$k].Age30
}
$teamClosureRate = [math]::Round($teamClosed / $teamTotal * 100, 1)

# ?? ROLLING 7-DAY HISTORY ??????????????????????????????????????????
$historyFile = "$OutputDir\backlog_history.json"
$history = @()
if (Test-Path $historyFile) {
    try { $history = Get-Content $historyFile | ConvertFrom-Json } catch {}
}

# Build today's snapshot
$todaySnap = [PSCustomObject]@{
    Date   = $ReportDate
    Agents = @{}
}
foreach ($name in $agentStats.Keys) {
    $todaySnap.Agents[$name] = @{ Active = $agentStats[$name].Active; Age30 = $agentStats[$name].Age30 }
}

# Replace today's entry if it already exists, otherwise append; keep last 5
$history = @($history | Where-Object { $_.Date -ne $ReportDate }) + @($todaySnap) | Select-Object -Last 5
$history | ConvertTo-Json -Depth 5 | Out-File $historyFile -Encoding UTF8

# Helper: build sparkline SVG for a metric across history entries
function Get-Sparkline($name, $metric) {
    $vals  = @()
    $dates = @()
    foreach ($h in $history) {
        $v = 0
        if ($h.Agents -and $h.Agents.PSObject.Properties[$name]) {
            $v = $h.Agents.PSObject.Properties[$name].Value.$metric
        } elseif ($h.Agents -and $h.Agents[$name]) {
            $v = $h.Agents[$name].$metric
        }
        $vals  += [int]$v
        $dates += $h.Date
    }
    if ($vals.Count -lt 2) { return "<span style='font-size:9px;color:#94a3b8;'>--</span>" }
    $max   = ($vals | Measure-Object -Maximum).Maximum
    $min   = ($vals | Measure-Object -Minimum).Minimum
    $range = [math]::Max($max - $min, 1)
    $w = 6; $gap = 2; $svgH = 20; $totalW = ($vals.Count * ($w + $gap)) - $gap
    $bars = ""
    $ci = 0
    foreach ($v in $vals) {
        $barH   = [math]::Max([math]::Round(($v - $min) / $range * ($svgH - 4) + 2), 2)
        $y      = $svgH - $barH
        $x      = $ci * ($w + $gap)
        $isLast = $ci -eq ($vals.Count - 1)
        $fill   = if ($isLast) { '#007DB8' } else { '#cbd5e1' }
        $label  = $dates[$ci] -replace ', 2026',''   # e.g. "June 18"
        $bars  += "<rect x='$x' y='$y' width='$w' height='$barH' rx='1' fill='$fill'><title>$label : $v</title></rect>"
        $ci++
    }
    return "<svg viewBox='0 0 $totalW $svgH' style='width:$(($vals.Count * 9))px;height:18px;vertical-align:middle;cursor:default;'>$bars</svg>"
}

# Helper: delta badge vs yesterday
function Get-Delta($name, $metric) {
    if ($history.Count -lt 2) { return "" }
    $prev = $history[$history.Count - 2]
    $prevVal = 0
    if ($prev.Agents -and $prev.Agents.PSObject.Properties[$name]) {
        $prevVal = $prev.Agents.PSObject.Properties[$name].Value.$metric
    } elseif ($prev.Agents -and $prev.Agents[$name]) {
        $prevVal = $prev.Agents[$name].$metric
    }
    $curr = $agentStats[$name].$metric
    $diff = $curr - $prevVal
    if ($diff -eq 0) { return "" }
    $col = if ($metric -eq 'Active' -or $metric -eq 'Age30') {
        if ($diff -gt 0) { '#e05252' } else { '#22c55e' }
    } else { '#64748b' }
    $sign = if ($diff -gt 0) { '+' } else { '' }
    return "<span style='font-size:10px;font-weight:700;color:$col;margin-left:4px;'>$sign$diff</span>"
}

function Get-TrendPanel($name, $rowId) {
    $rows = ""
    foreach ($h in $history) {
        $v = $null
        if ($h.Agents -and $h.Agents.PSObject.Properties[$name]) {
            $v = $h.Agents.PSObject.Properties[$name].Value
        } elseif ($h.Agents -and $h.Agents[$name]) {
            $v = $h.Agents[$name]
        }
        if (-not $v) { continue }
        $rows += "<tr><td style='padding:5px 12px;font-size:12px;color:#64748b;white-space:nowrap;'>$($h.Date)</td><td style='padding:5px 12px;text-align:center;font-size:12px;font-weight:700;color:#f59e0b;'>$($v.Active)</td><td style='padding:5px 12px;text-align:center;font-size:12px;font-weight:600;color:#1e293b;'>$($v.Age30)</td></tr>"
    }
    return @"
      <tr id="trend-$rowId" style="display:none;">
        <td colspan="6" style="padding:0 14px 12px 30px;border-bottom:1px solid #f1f5f9;background:#fafbfc;">
          <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:#94a3b8;margin-bottom:6px;padding-top:8px;">5-Run Trend</div>
          <table style="border-collapse:collapse;box-shadow:none;background:transparent;">
            <thead><tr>
              <th style="padding:4px 12px;font-size:10px;font-weight:600;color:#94a3b8;text-align:left;text-transform:uppercase;letter-spacing:0.5px;">Run Date</th>
              <th style="padding:4px 12px;font-size:10px;font-weight:600;color:#94a3b8;text-align:center;text-transform:uppercase;letter-spacing:0.5px;">Active</th>
              <th style="padding:4px 12px;font-size:10px;font-weight:600;color:#94a3b8;text-align:center;text-transform:uppercase;letter-spacing:0.5px;">30+ Days</th>
            </tr></thead>
            <tbody>$rows</tbody>
          </table>
        </td>
      </tr>
"@
}

$script:globalRowIndex = 0

function Build-TableRows($agentList) {
    $html = ""
    foreach ($name in $agentList) {
        if (-not $agentStats.ContainsKey($name)) { continue }
        $s   = $agentStats[$name]
        $cr  = if ($s.Total -gt 0) { [math]::Round($s.Closed / $s.Total * 100, 1) } else { 0 }
        $flag     = if ($s.Age30 -gt 0 -or $cr -lt 90) { '#e05252' } elseif ($s.Age16 -gt 0 -or $cr -lt 95) { '#f59e0b' } else { '#22c55e' }
        $crColor  = if ($cr -ge 95) { '#22c55e' } elseif ($cr -ge 90) { '#f59e0b' } else { '#e05252' }
        $a30Color = if ($s.Age30 -gt 0) { '#e05252' } else { '#1e293b' }
        $a16Color = if ($s.Age16 -gt 0) { '#f59e0b' } else { '#1e293b' }
        $deltaActive = Get-Delta $name 'Active'
        $deltaAge30  = Get-Delta $name 'Age30'
        $ri          = $script:globalRowIndex
        $trendPanel  = Get-TrendPanel $name $ri
        $script:globalRowIndex++
        $html += @"
      <tr style="cursor:pointer;" onclick="toggleTrend('trend-$ri', 'arr-$ri')">
        <td style="padding:10px 14px;border-bottom:1px solid #f1f5f9;">
          <div style="display:flex;align-items:center;gap:8px;">
            <span style="width:8px;height:8px;border-radius:50%;background:$flag;flex-shrink:0;display:inline-block;"></span>
            <span style="font-size:13px;font-weight:600;color:#1e293b;">$name</span>
            <span id="arr-$ri" style="font-size:11px;color:#94a3b8;margin-left:2px;">+</span>
          </div>
        </td>
        <td style="padding:10px 14px;border-bottom:1px solid #f1f5f9;text-align:center;font-size:13px;font-weight:700;color:#007DB8;">$($s.Total)</td>
        <td style="padding:10px 14px;border-bottom:1px solid #f1f5f9;text-align:center;font-size:13px;font-weight:600;color:$a16Color;">$($s.Age16)</td>
        <td style="padding:10px 14px;border-bottom:1px solid #f1f5f9;text-align:center;">
          <span style="font-size:13px;font-weight:700;color:$a30Color;">$($s.Age30)</span>$deltaAge30
        </td>
        <td style="padding:10px 14px;border-bottom:1px solid #f1f5f9;text-align:center;">
          <span style="font-size:13px;font-weight:700;color:#f59e0b;">$($s.Active)</span>$deltaActive
        </td>
        <td style="padding:10px 14px;border-bottom:1px solid #f1f5f9;text-align:center;font-size:13px;font-weight:700;color:$crColor;">$cr%</td>
      </tr>
      $trendPanel
"@
    }
    return $html
}

$phRowsHtml = Build-TableRows $phAgents
$deRowsHtml = Build-TableRows $deAgents

$outputFile = "$OutputDir\GTS_Agent_Backlog.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>GTS ESM - Agent Backlog Snapshot</title>
  <style>
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:'Segoe UI',Arial,sans-serif; background:#f1f5f9; color:#1e293b; padding:32px; max-width:960px; margin:0 auto; }
    .page-header { margin-bottom:24px; padding-bottom:16px; border-bottom:3px solid #007DB8; display:flex; justify-content:space-between; align-items:flex-end; }
    .page-header h1 { font-size:22px; font-weight:700; color:#007DB8; }
    .page-header .subtitle { font-size:12px; color:#64748b; margin-top:4px; }
    .page-header .datestamp { font-size:11px; color:#94a3b8; text-align:right; line-height:1.7; }
    .summary-strip { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:24px; }
    .summary-card { background:#fff; border-radius:12px; padding:14px 18px; box-shadow:0 1px 6px rgba(0,0,0,0.07); border-top:4px solid var(--accent); text-align:center; }
    .summary-card.blue  { --accent:#007DB8; } .summary-card.amber { --accent:#f59e0b; }
    .summary-card.red   { --accent:#e05252; } .summary-card.green { --accent:#22c55e; }
    .summary-label { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#94a3b8; margin-bottom:6px; }
    .summary-value { font-size:28px; font-weight:700; color:var(--accent); line-height:1; }
    .team-section { margin-bottom:24px; }
    .team-label { font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1.2px; color:#007DB8; margin-bottom:8px; padding-left:2px; }
    table { width:100%; border-collapse:collapse; background:#fff; border-radius:12px; overflow:hidden; box-shadow:0 1px 6px rgba(0,0,0,0.07); }
    thead tr { background:#f8fafc; }
    th { padding:10px 14px; font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:0.8px; color:#94a3b8; text-align:center; border-bottom:2px solid #e2e8f0; }
    th:first-child { text-align:left; }
    tbody tr:last-child td { border-bottom:none; }
    tbody tr:hover { background:#f8fafc; }
    .legend { display:flex; gap:16px; margin-bottom:20px; font-size:11px; color:#64748b; }
    .legend-dot { width:8px; height:8px; border-radius:50%; display:inline-block; margin-right:4px; }
    .page-footer { margin-top:12px; font-size:11px; color:#cbd5e1; text-align:center; padding-top:14px; border-top:1px solid #e2e8f0; }
  </style>
  <script>
    function toggleTrend(rowId, arrId) {
      var row = document.getElementById(rowId);
      var arr = document.getElementById(arrId);
      if (row.style.display === 'none') {
        row.style.display = 'table-row';
        arr.textContent = '-';
        arr.style.color = '#007DB8';
      } else {
        row.style.display = 'none';
        arr.textContent = '+';
        arr.style.color = '#94a3b8';
      }
    }
  </script>
</head>
<body>

  <div class="page-header">
    <div>
      <h1>GTS ESM - Agent Backlog Snapshot</h1>
      <div class="subtitle">Global Training Support &nbsp;.&nbsp; Case Backlog &amp; Aging Overview</div>
    </div>
    <div class="datestamp">As of $ReportDate<br>Source: SERVICE_CLOUD</div>
  </div>

  <div class="summary-strip">
    <div class="summary-card blue">
      <div class="summary-label">Team YTD Volume</div>
      <div class="summary-value">$($teamTotal.ToString('N0'))</div>
    </div>
    <div class="summary-card amber">
      <div class="summary-label">Total Active</div>
      <div class="summary-value">$($teamActive.ToString('N0'))</div>
    </div>
    <div class="summary-card red">
      <div class="summary-label">Cases 30+ Days</div>
      <div class="summary-value">$($teamAge30.ToString('N0'))</div>
    </div>
    <div class="summary-card green">
      <div class="summary-label">Team Closure Rate</div>
      <div class="summary-value">$teamClosureRate%</div>
    </div>
  </div>

  <div class="legend">
    <span><span class="legend-dot" style="background:#e05252;"></span> Red - 30+ day cases or closure rate below 90%</span>
    <span><span class="legend-dot" style="background:#f59e0b;"></span> Amber - 16-30 day cases or closure rate below 95%</span>
    <span><span class="legend-dot" style="background:#22c55e;"></span> Green - On track</span>
  </div>

  <div class="team-section">
    <div class="team-label">Philippines Team</div>
    <table>
      <thead>
        <tr>
          <th style="text-align:left;width:220px;">Agent</th>
          <th>YTD Volume</th>
          <th>16-30 Days</th>
          <th>30+ Days</th>
          <th>Active</th>
          <th>Closure Rate</th>
        </tr>
      </thead>
      <tbody>
        $phRowsHtml
      </tbody>
    </table>
  </div>

  <div class="team-section">
    <div class="team-label">Germany Team</div>
    <table>
      <thead>
        <tr>
          <th style="text-align:left;width:220px;">Agent</th>
          <th>YTD Volume</th>
          <th>16-30 Days</th>
          <th>30+ Days</th>
          <th>Active</th>
          <th>Closure Rate</th>
        </tr>
      </thead>
      <tbody>
        $deRowsHtml
      </tbody>
    </table>
  </div>

  <div class="page-footer">
    Generated $ReportDate &nbsp;.&nbsp; GTS ESM Operations &nbsp;.&nbsp; $(Split-Path $CasesXlsx -Leaf)
  </div>

</body>
</html>
"@

[System.IO.File]::WriteAllText($outputFile, $html, [System.Text.UTF8Encoding]::new($false))
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
