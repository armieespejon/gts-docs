# ============================================================
# GTS Coaching Log - Word Doc Generator (HTML-as-doc)
# Reads: Cases xlsx + CoachingLog_Entries.csv
# Outputs: One .doc per agent in Coaching Logs folder
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_CoachingLog_Word.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$DownloadsDir  = "C:\Users\I535893\Downloads"
$EntriesCsv    = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\GTS_CoachingLog_Entries.csv"
$OutputDir     = "C:\Users\I535893\Downloads\Claude and Me\Coaching Logs"
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

# Create output folder if needed
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

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

# ── LOAD DATA ────────────────────────────────────────────────
Write-Host "Loading data..." -ForegroundColor Yellow
$cHdr  = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE','EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13','L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY','CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

$sHdr    = @('MONTH','COL2','COL3','COL4','RATING','L1_CATEGORY','COMMENT1','COMMENT2','COL9','COL10')
$surveys = Import-Csv $SurveysCsv -Header $sHdr | Select-Object -Skip 2 | Where-Object { $_.RATING -match '^\d' }

$entries = @()
if (Test-Path $EntriesCsv) { $entries = Import-Csv $EntriesCsv }

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

$nameMap = @{
    'Nobela, Monica Frances'   = 'Monica Frances Nobela'
    'Quinto, Marviely'         = 'Marviely Quinto'
    'Chua, Sam'                = 'Sam Tomi Chua'
    'Santos, Ana Lizelle'      = 'Ana Lizelle Santos'
    'Tulagan, Nicole'          = 'Nicole Tulagan'
    'Mendoza, Edrian'          = 'Edrian Mendoza'
    'Cruz, Ann'                = 'Ann Cruz'
    'Valdez, Lemuel Von'       = 'Lemuel Von Valdez'
    'Sinag, Rendel Edison'     = 'Rendel Edison Sinag'
    'Evidente, Marie Cristine' = 'Marie Cristine Evidente'
    'Buenafe, Paul'            = 'Shaun Paul Buenafe'
    'Cada, Saldy'              = 'Saldy Cada'
    'Paulino, Justin'          = 'Justin Jan Paulino'
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

# ── COMPUTE STATS ────────────────────────────────────────────
Write-Host "Computing stats..." -ForegroundColor Yellow

$agentData = @{}
foreach ($a in $teamRoster) { $agentData[$a.Name] = @{ YTDIn=0; YTDClosed=0; Active=@(); Team=$a.Team } }

foreach ($row in $cases) {
    $name = Resolve-Name $row
    if ($null -eq $name -or -not $agentData.ContainsKey($name)) { continue }
    $agentData[$name].YTDIn++
    if ($row.STATUS -in @('Closed','Completed')) { $agentData[$name].YTDClosed++ }
    else { $agentData[$name].Active += $row }
}

$teamT2B   = ($surveys | Where-Object { [int]$_.RATING -ge 4 }).Count
$teamTotal = $surveys.Count
$teamCSAT  = if ($teamTotal -gt 0) { [math]::Round($teamT2B / $teamTotal * 100) } else { 0 }

# ── GENERATE ONE DOC PER AGENT ───────────────────────────────
Write-Host "Generating Word docs..." -ForegroundColor Yellow

foreach ($a in $teamRoster) {
    $name = $a.Name
    $d    = $agentData[$name]
    $rate = if ($d.YTDIn -gt 0) { [math]::Round($d.YTDClosed / $d.YTDIn * 100) } else { 0 }
    $d07=0; $d815=0; $d1630=0; $d30p=0; $ages=@()
    foreach ($row in $d.Active) {
        try {
            $age = ($Today - [datetime]::Parse($row.CREATED_ON_DATE)).Days
            $ages += $age
            if ($age -le 7) { $d07++ } elseif ($age -le 15) { $d815++ } elseif ($age -le 30) { $d1630++ } else { $d30p++ }
        } catch {}
    }
    $avgAge = if ($ages.Count -gt 0) { [math]::Round(($ages | Measure-Object -Average).Average,1) } else { 0 }

    $rateColor = if ($rate -ge 96) { '#047857' } elseif ($rate -ge 90) { '#b45309' } else { '#b91c1c' }
    $avgColor  = if ($avgAge -gt 20) { '#b91c1c' } elseif ($avgAge -gt 12) { '#b45309' } else { '#1d4ed8' }
    $csatColor = if ($teamCSAT -ge $TargetCSAT) { '#047857' } elseif ($teamCSAT -ge 70) { '#b45309' } else { '#b91c1c' }

    # Coaching entries for this agent (most recent first)
    $agentEntries = $entries | Where-Object { $_.Agent -eq $name } | Sort-Object { [datetime]::Parse($_.Date) } -Descending

    # Build entries HTML
    $entriesHtml = ""
    if ($agentEntries.Count -eq 0) {
        $entriesHtml = "<p style='color:#94a3b8;font-size:10pt;font-style:italic;margin:8px 0;'>No coaching entries yet.</p>"
    } else {
        foreach ($e in $agentEntries) {
            $statusColor = switch ($e.Status) {
                'Open'        { '#b91c1c' }
                'In Progress' { '#b45309' }
                default       { '#047857' }
            }
            $entriesHtml += @"
<table width='100%' cellpadding='0' cellspacing='0' style='margin-bottom:14px;border:1px solid #e2e8f0;border-radius:4px;'>
  <tr style='background:#f1f5f9;'>
    <td style='padding:7px 10px;font-size:11pt;font-weight:bold;color:#007DB8;width:30%;'>$($e.Date)</td>
    <td style='padding:7px 10px;font-size:10pt;font-weight:bold;color:#1e3a5f;text-align:center;width:45%;'>$($e.Topic)</td>
    <td style='padding:7px 10px;font-size:10pt;font-weight:bold;color:$statusColor;text-align:right;width:25%;'>$($e.Status.ToUpper())</td>
  </tr>
  <tr><td colspan='3' style='padding:10px 12px;'>
    <p style='font-size:8.5pt;color:#64748b;margin:0 0 8px 0;'>Snapshot at session: Active $($d.Active.Count) &nbsp;/&nbsp; 30+ days: $d30p &nbsp;/&nbsp; Closure: $rate% &nbsp;/&nbsp; CSAT: $teamCSAT%</p>
    <p style='font-size:9pt;font-weight:bold;color:#1e3a5f;margin:0 0 3px 0;'>Discussion</p>
    <p style='font-size:10pt;color:#374151;margin:0 0 10px 0;line-height:1.5;'>$($e.Discussion)</p>
    <table width='100%' cellpadding='8' cellspacing='0' style='background:#fffbeb;border-left:3px solid #b45309;'>
      <tr><td>
        <p style='font-size:8.5pt;font-weight:bold;color:#b45309;text-transform:uppercase;margin:0 0 4px 0;'>Agreed Action</p>
        <p style='font-size:10pt;color:#1e293b;margin:0 0 4px 0;'>$($e.Action)</p>
        <p style='font-size:10pt;color:#64748b;margin:0;'><strong>Due Date:</strong> &nbsp;$($e.DueDate)</p>
      </td></tr>
    </table>
  </td></tr>
</table>
"@
        }
    }

    # New entry template
    $newEntryHtml = @"
<table width='100%' cellpadding='0' cellspacing='0' style='margin-bottom:8px;border:1px solid #e2e8f0;background:#f8fafc;'>
  <tr>
    <td style='padding:7px 10px;font-size:10pt;color:#94a3b8;width:33%;'>Date: ___________</td>
    <td style='padding:7px 10px;font-size:10pt;color:#94a3b8;text-align:center;width:44%;'>Topic: ___________</td>
    <td style='padding:7px 10px;font-size:10pt;color:#94a3b8;text-align:right;width:23%;'>Status: ___________</td>
  </tr>
</table>
<p style='font-size:8.5pt;color:#94a3b8;margin:4px 0 12px 0;'>Snapshot: [auto-filled on next script run]</p>
"@

    $safeFileName = $name -replace '[\\/:*?"<>|]', '_'
    $outputFile   = "$OutputDir\CoachingLog_$safeFileName.doc"

    $html = @"
<html xmlns:o='urn:schemas-microsoft-com:office:office'
      xmlns:w='urn:schemas-microsoft-com:office:word'
      xmlns='http://www.w3.org/TR/REC-html40'>
<head>
<meta charset='UTF-8'/>
<xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml>
<style>
  @page { size:21cm 29.7cm; margin:1.8cm 2cm; }
  body { font-family:'Segoe UI',Arial,sans-serif; font-size:11pt; color:#1e293b; }
  table { border-collapse:collapse; }
  .sec-lbl { font-size:8pt; font-weight:bold; text-transform:uppercase; letter-spacing:1px; color:#64748b; border-bottom:1px solid #e2e8f0; margin:16px 0 8px 0; padding-bottom:4px; }
  .note-lbl { font-size:9pt; font-weight:bold; color:#1e3a5f; margin:12px 0 4px 0; }
  .note-line { border:none; border-bottom:1px solid #cbd5e1; margin:0 0 10px 0; height:18px; display:block; }
  hr { border:none; border-top:2px solid #e2e8f0; margin:8px 0 12px 0; }
</style>
</head>
<body>

<p style='font-size:22pt;font-weight:bold;color:#1e3a5f;margin:0 0 4px 0;'>GTS ESM &mdash; Manager Coaching Log</p>
<p style='font-size:13pt;color:#007DB8;margin:0 0 3px 0;'>$name &nbsp;&mdash;&nbsp; $($a.Team) Team</p>
<p style='font-size:9pt;color:#64748b;margin:0 0 10px 0;'>Last updated: $ReportDate &nbsp;&mdash;&nbsp; GTS Management &nbsp;&mdash;&nbsp; Confidential</p>
<hr/>

<div class='sec-lbl'>Current Performance Snapshot &mdash; as of $ReportDate</div>
<table width='100%' cellpadding='0' cellspacing='0' style='margin-bottom:12px;'>
  <tr>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:20%;'>YTD Volume</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:20%;'>Active Cases</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:20%;'>Closure Rate</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:20%;'>CSAT</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:20%;'>Avg Age</td>
  </tr>
  <tr>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:#1e3a5f;padding:8px;border:1px solid #e2e8f0;'>$($d.YTDIn)</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:#b45309;padding:8px;border:1px solid #e2e8f0;'>$($d.Active.Count)</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:$rateColor;padding:8px;border:1px solid #e2e8f0;'>$rate%</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:$csatColor;padding:8px;border:1px solid #e2e8f0;'>$teamCSAT%</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:$avgColor;padding:8px;border:1px solid #e2e8f0;'>$avgAge d</td>
  </tr>
</table>

<div class='sec-lbl'>Active Cases by Age</div>
<table width='100%' cellpadding='0' cellspacing='0' style='margin-bottom:14px;'>
  <tr>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:25%;'>0 - 7 Days</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:25%;'>8 - 15 Days</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:25%;'>16 - 30 Days</td>
    <td align='center' style='background:#1e3a5f;color:#fff;font-size:8pt;font-weight:bold;padding:6px;border:1px solid #fff;width:25%;'>30+ Days</td>
  </tr>
  <tr>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:#1e3a5f;padding:8px;border:1px solid #e2e8f0;'>$d07</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:#1d4ed8;padding:8px;border:1px solid #e2e8f0;'>$d815</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:#b45309;padding:8px;border:1px solid #e2e8f0;'>$d1630</td>
    <td align='center' style='background:#f8fafc;font-size:20pt;font-weight:bold;color:#b91c1c;padding:8px;border:1px solid #e2e8f0;'>$d30p</td>
  </tr>
</table>

<div class='sec-lbl'>Coaching Log &mdash; Most Recent First</div>
$entriesHtml

<div class='sec-lbl'>New Entry &mdash; [Next 1:1 Date]</div>
$newEntryHtml

<div class='note-lbl'>Discussion</div>
<span class='note-line'></span><span class='note-line'></span><span class='note-line'></span><span class='note-line'></span>

<div class='note-lbl'>Agreed Action</div>
<span class='note-line'></span><span class='note-line'></span>

<div class='note-lbl'>Due Date</div>
<span class='note-line'></span>

<div class='note-lbl'>Follow-up / Next Steps</div>
<span class='note-line'></span><span class='note-line'></span>

<p style='font-size:8pt;color:#94a3b8;text-align:center;margin-top:20px;border-top:1px solid #e2e8f0;padding-top:8px;'>
  GTS Management &nbsp;&mdash;&nbsp; Coaching Log &nbsp;&mdash;&nbsp; $name &nbsp;&mdash;&nbsp; Confidential
</p>

</body></html>
"@

    $html | Out-File $outputFile -Encoding UTF8
    Write-Host "  Generated: $safeFileName" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Done! $($teamRoster.Count) docs saved to:" -ForegroundColor Green
Write-Host "  $OutputDir" -ForegroundColor Yellow
Start-Process $OutputDir
