# ============================================================
# GTS Coaching Log - HTML Dashboard Generator
# Reads: Cases xlsx + CoachingLog_Entries.csv
# Outputs: GTS_CoachingLog.html (self-contained, dark mode)
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_CoachingLog_HTML.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$DownloadsDir  = "C:\Users\I535893\Downloads"
$EntriesCsv    = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\GTS_CoachingLog_Entries.csv"
$OutputDir     = "C:\Users\I535893\Downloads\Claude and Me"
$TargetClosure = 90
$TargetCSAT    = 80
# -- END CONFIG -----------------------------------------------

# Auto-detect latest Cases and Surveys xlsx files
$CasesXlsx = Get-ChildItem "$DownloadsDir\GTS ESM Cases_*.xlsx" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
$SurveysXlsx = Get-ChildItem "$DownloadsDir\GTS ESM Surveys_*.xlsx" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName

if (-not $CasesXlsx)   { Write-Host "ERROR: No GTS ESM Cases xlsx found in Downloads." -ForegroundColor Red; exit }
if (-not $SurveysXlsx) { Write-Host "ERROR: No GTS ESM Surveys xlsx found in Downloads." -ForegroundColor Red; exit }

Write-Host "  Cases file   : $([System.IO.Path]::GetFileName($CasesXlsx))" -ForegroundColor Gray
Write-Host "  Surveys file : $([System.IO.Path]::GetFileName($SurveysXlsx))" -ForegroundColor Gray

$Today      = Get-Date
$ReportDate = $Today.ToString("MMMM d, yyyy")
$Today      = [datetime]$Today.ToString("yyyy-MM-dd")

$CasesCsv   = "$env:TEMP\gts_coaching_cases.csv"
$SurveysCsv = "$env:TEMP\gts_coaching_surveys.csv"

# ── EXPORT XLSX TO CSV ───────────────────────────────────────
function Export-XlsxToCsv($xlsxPath, $csvPath) {
    Write-Host "  Exporting: $([System.IO.Path]::GetFileName($xlsxPath))" -ForegroundColor Cyan
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

# ── LOAD CASES ───────────────────────────────────────────────
Write-Host "Loading cases..." -ForegroundColor Yellow
$cHdr  = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE','EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13','L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY','CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

# ── LOAD SURVEYS ─────────────────────────────────────────────
Write-Host "Loading surveys..." -ForegroundColor Yellow
$sHdr    = @('MONTH','COL2','COL3','COL4','RATING','L1_CATEGORY','COMMENT1','COMMENT2','COL9','COL10')
$surveys = Import-Csv $SurveysCsv -Header $sHdr | Select-Object -Skip 2 | Where-Object { $_.RATING -match '^\d' }

# ── LOAD COACHING ENTRIES ────────────────────────────────────
Write-Host "Loading coaching entries..." -ForegroundColor Yellow
$entries = @()
if (Test-Path $EntriesCsv) {
    $entries = Import-Csv $EntriesCsv
}

# ── AGENT ROSTER ─────────────────────────────────────────────
$teamRoster = @(
    @{Name='Marviely Quinto';         Team='PH'},
    @{Name='Justin Jan Paulino';      Team='PH'},
    @{Name='Lemuel Von Valdez';       Team='PH'},
    @{Name='Monica Frances Nobela';   Team='PH'},
    @{Name='Rendel Edison Sinag';     Team='PH'},
    @{Name='Edrian Mendoza';          Team='PH'},
    @{Name='Shaun Paul Buenafe';      Team='PH'},
    @{Name='Saldy Cada';              Team='PH'},
    @{Name='Ana Lizelle Santos';      Team='PH'},
    @{Name='Sam Tomi Chua';           Team='PH'},
    @{Name='Nicole Tulagan';          Team='PH'},
    @{Name='Marie Cristine Evidente'; Team='PH'},
    @{Name='Ann Cruz';                Team='PH'},
    @{Name='Mritunjay Kumar Singh';   Team='DE'},
    @{Name='Shubham Thakur';          Team='DE'},
    @{Name='Anjali Tripathi';         Team='DE'},
    @{Name='Swarna Gupta';            Team='DE'},
    @{Name='Kriti Bhalla';            Team='DE'},
    @{Name='Narendran Raja';          Team='DE'}
)

# Map processor names from CSV to roster names
$nameMap = @{
    'Nobela, Monica Frances'  = 'Monica Frances Nobela'
    'Quinto, Marviely'        = 'Marviely Quinto'
    'Chua, Sam'               = 'Sam Tomi Chua'
    'Santos, Ana Lizelle'     = 'Ana Lizelle Santos'
    'Tulagan, Nicole'         = 'Nicole Tulagan'
    'Mendoza, Edrian'         = 'Edrian Mendoza'
    'Cruz, Ann'               = 'Ann Cruz'
    'Valdez, Lemuel Von'      = 'Lemuel Von Valdez'
    'Sinag, Rendel Edison'    = 'Rendel Edison Sinag'
    'Evidente, Marie Cristine'= 'Marie Cristine Evidente'
    'Buenafe, Paul'           = 'Shaun Paul Buenafe'
    'Cada, Saldy'             = 'Saldy Cada'
    'Paulino, Justin'         = 'Justin Jan Paulino'
    'Mritunjay Kumar Singh'   = 'Mritunjay Kumar Singh'
    'Shubham Thakur'          = 'Shubham Thakur'
    'Anjali Tripathi'         = 'Anjali Tripathi'
    'Swarna Gupta'            = 'Swarna Gupta'
    'Kriti Bhalla'            = 'Kriti Bhalla'
    'Narendran Raja'          = 'Narendran Raja'
}

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
    if ($nameMap.ContainsKey($n)) { return $nameMap[$n] }
    return $n
}

# ── COMPUTE AGENT STATS ───────────────────────────────────────
Write-Host "Computing stats..." -ForegroundColor Yellow

$rosterNames = $teamRoster | ForEach-Object { $_.Name }
$agentData   = @{}
foreach ($a in $teamRoster) { $agentData[$a.Name] = @{ YTDIn=0; YTDClosed=0; Active=@(); Team=$a.Team } }

foreach ($row in $cases) {
    $name = Resolve-Name $row
    if ($null -eq $name -or -not $agentData.ContainsKey($name)) { continue }
    $agentData[$name].YTDIn++
    if ($row.STATUS -in @('Closed','Completed')) { $agentData[$name].YTDClosed++ }
    else { $agentData[$name].Active += $row }
}

# CSAT per agent (from surveys - matched by processor name where possible)
$csatData = @{}
foreach ($name in $rosterNames) { $csatData[$name] = @{ T2B=0; Total=0 } }
foreach ($row in $surveys) {
    $rating = [int]$row.RATING
    # We don't have per-agent survey data, so CSAT shown as team average placeholder
    # Individual CSAT wired in when survey data includes processor name
}
$teamT2B    = ($surveys | Where-Object { [int]$_.RATING -ge 4 }).Count
$teamTotal  = $surveys.Count
$teamCSAT   = if ($teamTotal -gt 0) { [math]::Round($teamT2B / $teamTotal * 100) } else { 0 }

$agentStats = @()
foreach ($a in $teamRoster) {
    $name = $a.Name
    $d    = $agentData[$name]
    $rate = if ($d.YTDIn -gt 0) { [math]::Round($d.YTDClosed / $d.YTDIn * 100) } else { 0 }
    $d07=0; $d815=0; $d1630=0; $d30p=0; $ages=@()
    foreach ($row in $d.Active) {
        try {
            $age = ($Today - [datetime]::Parse($row.CREATED_ON_DATE)).Days
            $ages += $age
            if     ($age -le 7)  { $d07++ }
            elseif ($age -le 15) { $d815++ }
            elseif ($age -le 30) { $d1630++ }
            else                 { $d30p++ }
        } catch {}
    }
    $avgAge = if ($ages.Count -gt 0) { [math]::Round(($ages | Measure-Object -Average).Average,1) } else { 0 }

    # Flag logic
    $flag = if ($d30p -gt 0 -or $rate -lt $TargetClosure) { 'red' }
            elseif ($d1630 -gt 0 -or $rate -lt 95) { 'amber' }
            else { 'green' }

    $agentStats += [PSCustomObject]@{
        Name=$name; Team=$a.Team; YTDIn=$d.YTDIn; YTDClosed=$d.YTDClosed
        Rate=$rate; Active=$d.Active.Count; D07=$d07; D815=$d815
        D1630=$d1630; D30p=$d30p; AvgAge=$avgAge; Flag=$flag
        CSAT=$teamCSAT
    }
}

# Team totals
$tIn     = ($agentStats | Measure-Object YTDIn  -Sum).Sum
$tClosed = ($agentStats | Measure-Object YTDClosed -Sum).Sum
$tActive = ($agentStats | Measure-Object Active  -Sum).Sum
$tRate   = if ($tIn -gt 0) { [math]::Round($tClosed / $tIn * 100) } else { 0 }
$t30p    = ($agentStats | Measure-Object D30p    -Sum).Sum

# Open actions count
$openActions = ($entries | Where-Object { $_.Status -in @('Open','In Progress') }).Count
$flaggedCount = ($agentStats | Where-Object { $_.Flag -eq 'red' }).Count

# ── BUILD HTML COMPONENTS ────────────────────────────────────
Write-Host "Building HTML..." -ForegroundColor Yellow

function Get-Initials($name) {
    $parts = $name.Split(' ')
    if ($parts.Count -ge 2) { return ($parts[0][0].ToString() + $parts[-1][0].ToString()).ToUpper() }
    return $name.Substring(0,2).ToUpper()
}

function Get-TopicClass($topic) {
    if ($topic -match 'Volume|Aging|Closure') { return 'topic-volume' }
    if ($topic -match 'CSAT')       { return 'topic-csat' }
    if ($topic -match 'Attendance') { return 'topic-attendance' }
    if ($topic -match 'Process')    { return 'topic-process' }
    return 'topic-other'
}

function Get-StatusClass($status) {
    if ($status -eq 'Open')        { return 'status-open' }
    if ($status -eq 'In Progress') { return 'status-progress' }
    return 'status-closed'
}

function Get-RateColor($rate) {
    if ($rate -ge 96) { return 'green' }
    if ($rate -ge 90) { return 'amber' }
    return 'red'
}

function Build-AgentCard($s) {
    $initials   = Get-Initials $s.Name
    $rateColor  = Get-RateColor $s.Rate
    $agentEntries = $entries | Where-Object { $_.Agent -eq $s.Name } | Sort-Object { [datetime]::Parse($_.Date) } -Descending

    # Badges
    $badges = ""
    if ($s.D30p -gt 0)   { $badges += "<span class='badge red'>$($s.D30p) cases 30+ days</span>" }
    if ($s.D1630 -gt 0)  { $badges += "<span class='badge amber'>$($s.D1630) cases 16-30d</span>" }
    $badges += "<span class='badge $rateColor'>Closure $($s.Rate)%</span>"
    if ($s.CSAT -gt 0)   {
        $csatColor = if ($s.CSAT -ge $TargetCSAT) { 'green' } elseif ($s.CSAT -ge 70) { 'amber' } else { 'red' }
        $badges += "<span class='badge $csatColor'>CSAT $($s.CSAT)%</span>"
    }
    $openCount = ($agentEntries | Where-Object { $_.Status -in @('Open','In Progress') }).Count
    if ($openCount -gt 0) { $badges += "<span class='badge amber'>$openCount open action$(if($openCount -gt 1){'s'})</span>" }

    # KPI boxes
    $kpiRate = "<span class='kv $rateColor'>$($s.Rate)%</span>"
    $kpiAvg  = if ($s.AvgAge -gt 20) { "<span class='kv red'>$($s.AvgAge)d</span>" } elseif ($s.AvgAge -gt 12) { "<span class='kv amber'>$($s.AvgAge)d</span>" } else { "<span class='kv blue'>$($s.AvgAge)d</span>" }

    # Coaching entries HTML
    $entriesHtml = ""
    if ($agentEntries.Count -eq 0) {
        $entriesHtml = "<div class='no-entries'>No coaching entries yet.</div>"
    } else {
        foreach ($e in $agentEntries) {
            $topicClass  = Get-TopicClass $e.Topic
            $statusClass = Get-StatusClass $e.Status
            $entriesHtml += @"
<div class='c-entry'>
  <div class='c-entry-hdr'>
    <span class='c-date'>$($e.Date)</span>
    <span class='c-topic $topicClass'>$($e.Topic)</span>
    <span class='c-status $statusClass'>$($e.Status)</span>
  </div>
  <div class='c-field-lbl'>Discussion</div>
  <div class='c-field-val'>$($e.Discussion)</div>
  <div class='c-action'>
    <div class='c-action-lbl'>Agreed Action</div>
    <div class='c-action-val'>$($e.Action)</div>
    <div class='c-due'>Due: <span>$($e.DueDate)</span></div>
  </div>
</div>
"@
        }
    }

    return @"
<div class='agent-card flag-$($s.Flag)' data-agent='$($s.Name)' data-flag='flag-$($s.Flag)' data-team='$($s.Team)'>
  <div class='card-hdr' onclick='toggle(this)'>
    <div class='avatar'>$initials</div>
    <div class='agent-info'>
      <div class='agent-name'>$($s.Name)</div>
      <div class='agent-team'>$($s.Team) Team</div>
      <div class='badge-row'>$badges</div>
    </div>
    <div class='chevron'>&#9660;</div>
  </div>
  <div class='card-body'>
    <div class='kpi-row'>
      <div class='kbox'><div class='kl'>YTD Volume</div><div class='kv blue'>$($s.YTDIn)</div></div>
      <div class='kbox'><div class='kl'>Active</div><div class='kv blue'>$($s.Active)</div></div>
      <div class='kbox'><div class='kl'>Closure Rate</div>$kpiRate</div>
      <div class='kbox'><div class='kl'>CSAT</div><div class='kv $(if($s.CSAT -ge $TargetCSAT){"green"}elseif($s.CSAT -ge 70){"amber"}else{"red"})'>$($s.CSAT)%</div></div>
      <div class='kbox'><div class='kl'>Avg Age</div>$kpiAvg</div>
    </div>
    <div class='age-row'>
      <div class='abox a07'><div class='av'>$($s.D07)</div><div class='al'>0-7 Days</div></div>
      <div class='abox a815'><div class='av'>$($s.D815)</div><div class='al'>8-15 Days</div></div>
      <div class='abox a16'><div class='av'>$($s.D1630)</div><div class='al'>16-30 Days</div></div>
      <div class='abox a30'><div class='av'>$($s.D30p)</div><div class='al'>30+ Days</div></div>
    </div>
    <div class='sec-lbl'>Coaching Log</div>
    $entriesHtml
  </div>
</div>
"@
}

$cardsHtml = ($agentStats | ForEach-Object { Build-AgentCard $_ }) -join "`n"

# ── WRITE HTML FILE ──────────────────────────────────────────
$outputFile = "$OutputDir\GTS_CoachingLog.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>GTS Coaching Log | $ReportDate</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Arial,sans-serif;background:#0d1117;color:#e6edf3;min-height:100vh}

    .page-hdr{background:linear-gradient(135deg,#007DB8,#005a8a);padding:22px 32px 18px;border-bottom:3px solid #F0AB00}
    .page-hdr h1{font-size:21px;font-weight:700;color:#fff;margin-bottom:3px}
    .page-hdr p{font-size:12px;color:#bde3f7}

    .team-strip{display:grid;grid-template-columns:repeat(6,1fr);gap:1px;background:#30363d;border-bottom:1px solid #30363d}
    .s-tile{background:#161b22;padding:13px 18px;text-align:center}
    .s-lbl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#8b949e;margin-bottom:5px}
    .s-val{font-size:24px;font-weight:700}
    .s-val.blue{color:#38b6e8}.s-val.green{color:#3fb950}.s-val.amber{color:#F0AB00}.s-val.red{color:#f85149}

    .filter-bar{background:#161b22;padding:12px 32px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #30363d}
    .filter-bar label{font-size:11px;color:#8b949e;font-weight:700;text-transform:uppercase;letter-spacing:.6px}
    .filter-bar select{background:#0d1117;border:1px solid #30363d;color:#e6edf3;padding:5px 10px;border-radius:6px;font-size:12px;outline:none}
    .filter-bar select:focus{border-color:#38b6e8}
    .f-count{margin-left:auto;font-size:12px;color:#8b949e}

    .grid{padding:18px 32px;display:grid;grid-template-columns:repeat(auto-fill,minmax(480px,1fr));gap:14px}

    .agent-card{background:#161b22;border:1px solid #30363d;border-radius:10px;overflow:hidden;transition:border-color .2s}
    .agent-card:hover{border-color:#38b6e8}
    .agent-card.flag-red{border-left:4px solid #f85149}
    .agent-card.flag-amber{border-left:4px solid #F0AB00}
    .agent-card.flag-green{border-left:4px solid #3fb950}

    .card-hdr{padding:13px 15px 10px;display:flex;align-items:center;gap:11px;cursor:pointer;user-select:none}
    .avatar{width:38px;height:38px;border-radius:50%;background:#007DB8;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:#fff;flex-shrink:0}
    .agent-info{flex:1}
    .agent-name{font-size:14px;font-weight:700;color:#e6edf3}
    .agent-team{font-size:11px;color:#8b949e;margin-top:1px}
    .badge-row{display:flex;gap:5px;margin-top:5px;flex-wrap:wrap}
    .badge{font-size:10px;font-weight:700;padding:2px 7px;border-radius:20px;text-transform:uppercase;letter-spacing:.4px}
    .badge.green{background:#0d2818;color:#3fb950;border:1px solid #1a4a2e}
    .badge.amber{background:#2d1f00;color:#F0AB00;border:1px solid #4a3500}
    .badge.red{background:#2d0f0f;color:#f85149;border:1px solid #4a1a1a}
    .badge.blue{background:#0d2030;color:#38b6e8;border:1px solid #1a3a50}
    .chevron{font-size:16px;color:#8b949e;transition:transform .2s}
    .agent-card.open .chevron{transform:rotate(180deg)}

    .card-body{display:none;padding:0 15px 15px}
    .agent-card.open .card-body{display:block}

    .kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:7px;margin-bottom:12px;padding-top:4px}
    .kbox{background:#0d1117;border-radius:7px;padding:9px 7px;text-align:center;border:1px solid #21262d}
    .kl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#8b949e;margin-bottom:4px}
    .kv{font-size:20px;font-weight:700;display:block}
    .kv.blue{color:#38b6e8}.kv.green{color:#3fb950}.kv.amber{color:#F0AB00}.kv.red{color:#f85149}

    .age-row{display:grid;grid-template-columns:repeat(4,1fr);gap:7px;margin-bottom:13px}
    .abox{border-radius:7px;padding:8px 9px;text-align:center}
    .abox.a07{background:#0d2030;border:1px solid #1a3a50}.abox.a815{background:#0d1a30;border:1px solid #1a2d50}
    .abox.a16{background:#2d1f00;border:1px solid #4a3500}.abox.a30{background:#2d0f0f;border:1px solid #4a1a1a}
    .av{font-size:22px;font-weight:700;line-height:1}
    .abox.a07 .av{color:#38b6e8}.abox.a815 .av{color:#79b8f5}.abox.a16 .av{color:#F0AB00}.abox.a30 .av{color:#f85149}
    .al{font-size:9px;color:#8b949e;margin-top:3px;font-weight:600}

    .sec-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#8b949e;margin:13px 0 8px;padding-bottom:5px;border-bottom:1px solid #21262d}

    .c-entry{background:#0d1117;border:1px solid #21262d;border-radius:7px;padding:11px 13px;margin-bottom:8px}
    .c-entry-hdr{display:flex;align-items:center;gap:7px;margin-bottom:8px;flex-wrap:wrap}
    .c-date{font-size:11px;font-weight:700;color:#38b6e8}
    .c-topic{font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px}
    .topic-volume{background:#0d2030;color:#38b6e8;border:1px solid #1a3a50}
    .topic-csat{background:#0d2818;color:#3fb950;border:1px solid #1a4a2e}
    .topic-attendance{background:#2d2000;color:#F0AB00;border:1px solid #4a3800}
    .topic-process{background:#2d1a2d;color:#d2a8ff;border:1px solid #4a2a4a}
    .topic-other{background:#1c2128;color:#8b949e;border:1px solid #30363d}
    .c-status{font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;margin-left:auto}
    .status-open{background:#2d0f0f;color:#f85149;border:1px solid #4a1a1a}
    .status-progress{background:#2d1f00;color:#F0AB00;border:1px solid #4a3500}
    .status-closed{background:#0d2818;color:#3fb950;border:1px solid #1a4a2e}
    .c-field-lbl{font-size:10px;color:#8b949e;font-weight:600;margin-bottom:2px}
    .c-field-val{font-size:12px;color:#c9d1d9;line-height:1.5;margin-bottom:8px}
    .c-action{background:#161b22;border-left:3px solid #F0AB00;padding:6px 10px;border-radius:0 5px 5px 0}
    .c-action-lbl{font-size:9px;color:#F0AB00;font-weight:700;text-transform:uppercase;letter-spacing:.6px}
    .c-action-val{font-size:12px;color:#c9d1d9;margin-top:2px}
    .c-due{font-size:10px;color:#8b949e;margin-top:3px}
    .c-due span{color:#F0AB00;font-weight:600}
    .no-entries{text-align:center;padding:16px;color:#8b949e;font-size:12px;font-style:italic}

    .page-footer{text-align:center;padding:18px;font-size:11px;color:#8b949e;border-top:1px solid #21262d;margin-top:8px}
  </style>
</head>
<body>

<div class="page-hdr">
  <h1>GTS ESM &mdash; Manager Coaching Log</h1>
  <p>$ReportDate &nbsp;|&nbsp; GTS Management &nbsp;|&nbsp; Confidential &nbsp;|&nbsp; 19 Agents</p>
</div>

<div class="team-strip">
  <div class="s-tile"><div class="s-lbl">Total Active</div><div class="s-val blue">$tActive</div></div>
  <div class="s-tile"><div class="s-lbl">YTD Closure Rate</div><div class="s-val $(if($tRate -ge 90){'amber'}else{'red'})">$tRate%</div></div>
  <div class="s-tile"><div class="s-lbl">Team CSAT</div><div class="s-val $(if($teamCSAT -ge $TargetCSAT){'green'}elseif($teamCSAT -ge 70){'amber'}else{'red'})">$teamCSAT%</div></div>
  <div class="s-tile"><div class="s-lbl">30+ Day Cases</div><div class="s-val $(if($t30p -gt 0){'red'}else{'green'})">$t30p</div></div>
  <div class="s-tile"><div class="s-lbl">Open Actions</div><div class="s-val $(if($openActions -gt 0){'amber'}else{'green'})">$openActions</div></div>
  <div class="s-tile"><div class="s-lbl">Agents Flagged</div><div class="s-val $(if($flaggedCount -gt 0){'red'}else{'green'})">$flaggedCount</div></div>
</div>

<div class="filter-bar">
  <label>Agent</label>
  <select id="fAgent" onchange="applyFilters()">
    <option value="">All Agents</option>
    $(($agentStats | ForEach-Object { "<option>$($_.Name)</option>" }) -join "")
  </select>
  <label>Team</label>
  <select id="fTeam" onchange="applyFilters()">
    <option value="">All Teams</option>
    <option value="PH">PH Team</option>
    <option value="DE">DE Team</option>
  </select>
  <label>Topic</label>
  <select id="fTopic" onchange="applyFilters()">
    <option value="">All Topics</option>
    <option value="topic-volume">Case Volume / Aging</option>
    <option value="topic-csat">CSAT</option>
    <option value="topic-attendance">Attendance</option>
    <option value="topic-process">Process Gap</option>
    <option value="topic-other">Other</option>
  </select>
  <label>Status</label>
  <select id="fStatus" onchange="applyFilters()">
    <option value="">All Statuses</option>
    <option value="status-open">Open</option>
    <option value="status-progress">In Progress</option>
    <option value="status-closed">Closed</option>
  </select>
  <label>Flag</label>
  <select id="fFlag" onchange="applyFilters()">
    <option value="">All</option>
    <option value="flag-red">Needs Attention</option>
    <option value="flag-amber">Watch</option>
    <option value="flag-green">On Track</option>
  </select>
  <span class="f-count" id="fCount">Showing 19 of 19 agents</span>
</div>

<div class="grid" id="grid">
$cardsHtml
</div>

<div class="page-footer">GTS Management &nbsp;|&nbsp; Manager Coaching Log &nbsp;|&nbsp; $ReportDate &nbsp;|&nbsp; Confidential</div>

<script>
function toggle(hdr) { hdr.closest('.agent-card').classList.toggle('open'); }
function applyFilters() {
  const agent  = document.getElementById('fAgent').value;
  const team   = document.getElementById('fTeam').value;
  const topic  = document.getElementById('fTopic').value;
  const status = document.getElementById('fStatus').value;
  const flag   = document.getElementById('fFlag').value;
  const cards  = document.querySelectorAll('.agent-card');
  let visible  = 0;
  cards.forEach(card => {
    const aMatch = !agent  || card.dataset.agent === agent;
    const tMatch = !team   || card.dataset.team  === team;
    const fMatch = !flag   || card.dataset.flag  === flag;
    let topicMatch  = !topic;
    let statusMatch = !status;
    if (topic || status) {
      card.querySelectorAll('.c-entry').forEach(entry => {
        if (topic  && entry.querySelector('.' + topic))  topicMatch  = true;
        if (status && entry.querySelector('.' + status)) statusMatch = true;
      });
      if (card.querySelectorAll('.c-entry').length === 0) {
        if (topic)  topicMatch  = false;
        if (status) statusMatch = false;
      }
    }
    const show = aMatch && tMatch && fMatch && topicMatch && statusMatch;
    card.style.display = show ? '' : 'none';
    if (show) visible++;
  });
  document.getElementById('fCount').textContent = 'Showing ' + visible + ' of ' + cards.length + ' agents';
}
</script>
</body>
</html>
"@

$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! HTML saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
