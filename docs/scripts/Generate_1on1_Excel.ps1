# ============================================================
# GTS Team 1:1 Coaching Review -- Excel Generator
# Builds one Excel workbook: Manager Summary tab + one tab per agent
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File "Generate_1on1_Excel.ps1"
# ============================================================

# -- CONFIG ---------------------------------------------------
$CasesXlsx  = "C:\Users\I535893\Downloads\GTS ESM Cases_June17.xlsx"
$OutputDir  = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"
$ReportDate = "June 17, 2026"
$Today      = [datetime]'2026-06-17'
$TargetRate = 90
# -- END CONFIG -----------------------------------------------

$CasesCsv = "$env:TEMP\gts_cases_coaching.csv"

function Export-XlsxToCsv($xlsxPath, $csvPath) {
    Write-Host "  Exporting source data..." -ForegroundColor Cyan
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($xlsxPath)
    $wb.SaveAs($csvPath, 6)
    $wb.Close($false); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

Export-XlsxToCsv $CasesXlsx $CasesCsv

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE','EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13','L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY','CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing agent stats..." -ForegroundColor Yellow

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
foreach ($name in $teamNames) { $agentData[$name] = @{ YTDIn=0; YTDClosed=0; Active=@() } }

foreach ($row in $cases) {
    $name = Resolve-Name $row
    if ($null -eq $name -or -not $agentData.ContainsKey($name)) { continue }
    $agentData[$name].YTDIn++
    if ($row.STATUS -in @('Closed','Completed')) { $agentData[$name].YTDClosed++ }
    else { $agentData[$name].Active += $row }
}

$agentStats = @()
foreach ($name in $teamNames) {
    $a = $agentData[$name]
    $rate = if ($a.YTDIn -gt 0) { [math]::Round($a.YTDClosed / $a.YTDIn * 100) } else { 0 }
    $d07=0; $d815=0; $d1630=0; $d30p=0; $ages=@()
    foreach ($row in $a.Active) {
        try {
            $age = ($Today - [datetime]::Parse($row.CREATED_ON_DATE)).Days
            $ages += $age
            if ($age -le 7) { $d07++ } elseif ($age -le 15) { $d815++ } elseif ($age -le 30) { $d1630++ } else { $d30p++ }
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

$tIn     = ($agentStats | Measure-Object YTDIn -Sum).Sum
$tClosed = ($agentStats | Measure-Object YTDClosed -Sum).Sum
$tActive = ($agentStats | Measure-Object Active -Sum).Sum
$tRate   = [math]::Round($tClosed / $tIn * 100)
$t07     = ($agentStats | Measure-Object D07 -Sum).Sum
$t815    = ($agentStats | Measure-Object D815 -Sum).Sum
$t1630   = ($agentStats | Measure-Object D1630 -Sum).Sum
$t30p    = ($agentStats | Measure-Object D30p -Sum).Sum

Write-Host "Building Excel workbook..." -ForegroundColor Yellow

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

# ── Color constants ───────────────────────────────────────────
$navyBg    = 0x5F3A1E   # Excel uses BGR -- #1E3A5F
$navyBg    = [long]0x5F3A1E
$whiteFg   = 16777215
$lightBlue = [long]0xF6EFF0  # #F0EFF6 light lavender
$headerBg  = [long]0x5F3A1E

# Helper: RGB to Excel color (BGR)
function rgb($r,$g,$b) { return [long]($b * 65536 + $g * 256 + $r) }

$cNavy    = rgb 30  58  95
$cBlue    = rgb 29  78 216
$cGreen   = rgb  4 120  87
$cAmber   = rgb 217 119  6
$cRed     = rgb 185  28  28
$cLightBg = rgb 248 250 252
$cYellow  = rgb 254 243 199
$cPink    = rgb 254 226 226
$cGreenBg = rgb 220 252 231
$cWhite   = rgb 255 255 255
$cGray    = rgb 100 116 139
$cBorder  = rgb 226 232 240

function Set-Cell($ws, $row, $col, $val, $bold=$false, $size=11, $fg=-1, $bg=-1, $align="Left", $wrap=$false, $italic=$false) {
    $c = $ws.Cells($row, $col)
    $c.Value2 = $val
    $c.Font.Bold = $bold
    $c.Font.Size = $size
    $c.Font.Italic = $italic
    if ($fg -ge 0) { $c.Font.Color = $fg }
    if ($bg -ge 0) { $c.Interior.Color = $bg }
    $c.HorizontalAlignment = switch ($align) {
        "Center" { -4108 } "Right" { -4152 } default { -4131 }
    }
    $c.WrapText = $wrap
}

function Set-Border($ws, $row1, $col1, $row2, $col2) {
    $range = $ws.Range($ws.Cells($row1,$col1), $ws.Cells($row2,$col2))
    $range.Borders.LineStyle = 1
    $range.Borders.Color = $cBorder
    $range.Borders.Weight = 2
}

function Merge-And-Set($ws, $row, $col1, $col2, $val, $bold=$false, $size=11, $fg=-1, $bg=-1, $align="Left") {
    $range = $ws.Range($ws.Cells($row,$col1), $ws.Cells($row,$col2))
    $range.Merge() | Out-Null
    $range.Value2 = $val
    $range.Font.Bold = $bold
    $range.Font.Size = $size
    if ($fg -ge 0) { $range.Font.Color = $fg }
    if ($bg -ge 0) { $range.Interior.Color = $bg }
    $range.HorizontalAlignment = switch ($align) { "Center" { -4108 } "Right" { -4152 } default { -4131 } }
}

# ══════════════════════════════════════════════════════════════
# SHEET 1: MANAGER SUMMARY
# ══════════════════════════════════════════════════════════════
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Manager Summary"
$ws.Tab.Color = $cNavy

# Column widths
$ws.Columns(1).ColumnWidth = 28
$ws.Columns(2).ColumnWidth = 12
$ws.Columns(3).ColumnWidth = 12
$ws.Columns(4).ColumnWidth = 12
$ws.Columns(5).ColumnWidth = 10
$ws.Columns(6).ColumnWidth = 8
$ws.Columns(7).ColumnWidth = 8
$ws.Columns(8).ColumnWidth = 10
$ws.Columns(9).ColumnWidth = 8
$ws.Columns(10).ColumnWidth = 10
$ws.Columns(11).ColumnWidth = 10

# Title
$ws.Rows(1).RowHeight = 36
Merge-And-Set $ws 1 1 11 "GTS Team -- Manager Summary & 1:1 Review" $true 16 $cWhite $cNavy "Left"
$ws.Cells(1,1).IndentLevel = 1

$ws.Rows(2).RowHeight = 20
Merge-And-Set $ws 2 1 11 "As of $ReportDate  |  GTS Management  |  Confidential" $false 10 (rgb 147 197 253) $cNavy "Left"
$ws.Cells(2,1).IndentLevel = 1

# Team KPI tiles (row 4-6)
$ws.Rows(3).RowHeight = 10
$kpis = @(
    @{L="YTD Inflow"; V=$tIn.ToString("N0"); C=$cNavy}
    @{L="YTD Closures"; V=$tClosed.ToString("N0"); C=$cGreen}
    @{L="YTD Closure Rate"; V="$tRate%"; C=if($tRate -ge 93){$cGreen}elseif($tRate -ge 90){$cAmber}else{$cRed}}
    @{L="Active Cases"; V=$tActive.ToString("N0"); C=$cAmber}
)
$col = 1
foreach ($kpi in $kpis) {
    $ws.Rows(4).RowHeight = 16
    $ws.Rows(5).RowHeight = 30
    $ws.Rows(6).RowHeight = 14
    $c = $ws.Cells(4, $col); $c.Value2 = $kpi.L; $c.Font.Size = 9; $c.Font.Bold = $true; $c.Font.Color = $cGray; $c.Interior.Color = $cLightBg
    $c2 = $ws.Cells(5, $col); $c2.Value2 = $kpi.V; $c2.Font.Size = 22; $c2.Font.Bold = $true; $c2.Font.Color = $kpi.C; $c2.Interior.Color = $cLightBg; $c2.HorizontalAlignment = -4131; $c2.IndentLevel = 1
    $ws.Cells(6, $col).Interior.Color = $kpi.C
    $ws.Cells(6, $col).RowHeight = 4
    $col += 2
    if ($col -eq 9) { $col++ }
}

# Aging buckets (row 8-10)
$ws.Rows(7).RowHeight = 10
$agings = @(
    @{L="0-7 Days"; V=$t07; C=$cNavy; BG=(rgb 239 246 255)}
    @{L="8-15 Days"; V=$t815; C=$cBlue; BG=(rgb 219 234 254)}
    @{L="16-30 Days"; V=$t1630; C=$cAmber; BG=$cYellow}
    @{L=">30 Days"; V=$t30p; C=$cRed; BG=$cPink}
)
$col = 1
foreach ($ag in $agings) {
    $ws.Rows(8).RowHeight = 14
    $ws.Rows(9).RowHeight = 30
    $ws.Rows(10).RowHeight = 14
    $c = $ws.Cells(8, $col); $c.Value2 = $ag.L; $c.Font.Size = 9; $c.Font.Bold = $true; $c.Font.Color = $ag.C; $c.Interior.Color = $ag.BG; $c.HorizontalAlignment = -4108
    $c2 = $ws.Cells(9, $col); $c2.Value2 = $ag.V; $c2.Font.Size = 24; $c2.Font.Bold = $true; $c2.Font.Color = $ag.C; $c2.Interior.Color = $ag.BG; $c2.HorizontalAlignment = -4108
    $c3 = $ws.Cells(10, $col); $c3.Interior.Color = $ag.BG
    # Merge across 2 cols for aging
    $ws.Range($ws.Cells(8,$col), $ws.Cells(8,$col+1)).Merge() | Out-Null
    $ws.Range($ws.Cells(9,$col), $ws.Cells(9,$col+1)).Merge() | Out-Null
    $ws.Range($ws.Cells(10,$col), $ws.Cells(10,$col+1)).Merge() | Out-Null
    $col += 3
}

# Needs Attention header
$ws.Rows(12).RowHeight = 22
Merge-And-Set $ws 12 1 11 "NEEDS ATTENTION -- Agents with >30-day cases or closure rate below $TargetRate%" $true 11 $cRed $cPink "Left"
$ws.Cells(12,1).IndentLevel = 1

# Needs Attention table headers
$ws.Rows(13).RowHeight = 18
$hdrs = @("Agent","Active",">30d Cases","16-30d Cases","Closure Rate")
for ($i=0; $i -lt $hdrs.Count; $i++) {
    $c = $ws.Cells(13, $i+1); $c.Value2 = $hdrs[$i]; $c.Font.Bold = $true; $c.Font.Size = 10
    $c.Font.Color = $cWhite; $c.Interior.Color = $cNavy; $c.HorizontalAlignment = if($i -eq 0){-4131}else{-4108}
    if ($i -eq 0) { $c.IndentLevel = 1 }
}

$needsAttn = $agentStats | Where-Object { $_.D30p -gt 0 -or $_.Rate -lt $TargetRate } | Sort-Object { $_.D30p * 100 + $_.D1630 } -Descending
$r = 14
foreach ($s in $needsAttn) {
    $ws.Rows($r).RowHeight = 16
    $bg = if ($r % 2 -eq 0) { $cLightBg } else { $cWhite }
    $c = $ws.Cells($r,1); $c.Value2 = $s.Name; $c.Font.Bold = $true; $c.Font.Size = 10; $c.Interior.Color = $bg; $c.IndentLevel = 1
    $ws.Cells($r,2).Value2 = $s.Active; $ws.Cells($r,2).Interior.Color = $bg; $ws.Cells($r,2).HorizontalAlignment = -4108
    $c30 = $ws.Cells($r,3); $c30.Value2 = $s.D30p; $c30.HorizontalAlignment = -4108; $c30.Interior.Color = $bg
    if ($s.D30p -gt 0) { $c30.Font.Color = $cRed; $c30.Font.Bold = $true }
    $c16 = $ws.Cells($r,4); $c16.Value2 = $s.D1630; $c16.HorizontalAlignment = -4108; $c16.Interior.Color = $bg
    if ($s.D1630 -gt 0) { $c16.Font.Color = $cAmber; $c16.Font.Bold = $true }
    $cr = $ws.Cells($r,5); $cr.Value2 = "$($s.Rate)%"; $cr.HorizontalAlignment = -4108; $cr.Interior.Color = $bg; $cr.Font.Bold = $true
    $cr.Font.Color = if ($s.Rate -lt $TargetRate) { $cRed } else { $cBlue }
    $r++
}

# Performing Well
$r += 1
$ws.Rows($r).RowHeight = 22
Merge-And-Set $ws $r 1 11 "PERFORMING WELL -- Agents with 95%+ closure rate and no >30-day cases" $true 11 $cGreen $cGreenBg "Left"
$ws.Cells($r,1).IndentLevel = 1
$r++

$ws.Rows($r).RowHeight = 18
$hdrs2 = @("Agent","Active","Avg Age","Closure Rate")
for ($i=0; $i -lt $hdrs2.Count; $i++) {
    $c = $ws.Cells($r, $i+1); $c.Value2 = $hdrs2[$i]; $c.Font.Bold = $true; $c.Font.Size = 10
    $c.Font.Color = $cWhite; $c.Interior.Color = $cGreen; $c.HorizontalAlignment = if($i -eq 0){-4131}else{-4108}
    if ($i -eq 0) { $c.IndentLevel = 1 }
}
$r++

$performing = $agentStats | Where-Object { $_.D30p -eq 0 -and $_.Rate -ge 95 } | Sort-Object Rate -Descending
foreach ($s in $performing) {
    $ws.Rows($r).RowHeight = 16
    $bg = if ($r % 2 -eq 0) { $cLightBg } else { $cWhite }
    $c = $ws.Cells($r,1); $c.Value2 = $s.Name; $c.Font.Bold = $true; $c.Font.Color = $cGreen; $c.Font.Size = 10; $c.Interior.Color = $bg; $c.IndentLevel = 1
    $ws.Cells($r,2).Value2 = $s.Active; $ws.Cells($r,2).HorizontalAlignment = -4108; $ws.Cells($r,2).Interior.Color = $bg
    $ws.Cells($r,3).Value2 = $s.AvgAge; $ws.Cells($r,3).HorizontalAlignment = -4108; $ws.Cells($r,3).Interior.Color = $bg
    $cr = $ws.Cells($r,4); $cr.Value2 = "$($s.Rate)%"; $cr.Font.Color = $cGreen; $cr.Font.Bold = $true; $cr.HorizontalAlignment = -4108; $cr.Interior.Color = $bg
    $r++
}

# Full Team table
$r += 1
$ws.Rows($r).RowHeight = 22
Merge-And-Set $ws $r 1 11 "FULL TEAM OVERVIEW" $true 11 $cWhite $cNavy "Left"
$ws.Cells($r,1).IndentLevel = 1
$r++

$ws.Rows($r).RowHeight = 18
$fhdrs = @("Agent","YTD In","YTD Closed","YTD Rate","Active","0-7d","8-15d","16-30d",">30d","Avg Age","Max Age")
for ($i=0; $i -lt $fhdrs.Count; $i++) {
    $c = $ws.Cells($r, $i+1); $c.Value2 = $fhdrs[$i]; $c.Font.Bold = $true; $c.Font.Size = 10; $c.Font.Color = $cWhite; $c.Interior.Color = $cNavy
    $c.HorizontalAlignment = if($i -eq 0){-4131}else{-4108}
    if ($i -eq 0) { $c.IndentLevel = 1 }
}
$r++

foreach ($s in $agentStats) {
    $ws.Rows($r).RowHeight = 16
    $bg = if ($r % 2 -eq 0) { $cLightBg } else { $cWhite }
    $rateColor = if($s.Rate -ge 96){$cGreen}elseif($s.Rate -ge 93){$cBlue}elseif($s.Rate -ge 90){$cAmber}else{$cRed}
    $c = $ws.Cells($r,1); $c.Value2 = $s.Name; $c.Font.Size = 10; $c.Interior.Color = $bg; $c.IndentLevel = 1
    @(2,3) | ForEach-Object { $cc = $ws.Cells($r,$_); $cc.Value2 = $agentStats[($agentStats.IndexOf($s))].($fhdrs[$_-1] -replace ' ',''); $cc.HorizontalAlignment=-4108; $cc.Interior.Color=$bg }
    $ws.Cells($r,2).Value2 = $s.YTDIn; $ws.Cells($r,2).HorizontalAlignment=-4108; $ws.Cells($r,2).Interior.Color=$bg
    $ws.Cells($r,3).Value2 = $s.YTDClosed; $ws.Cells($r,3).HorizontalAlignment=-4108; $ws.Cells($r,3).Interior.Color=$bg
    $ws.Cells($r,4).Value2 = "$($s.Rate)%"; $ws.Cells($r,4).Font.Color=$rateColor; $ws.Cells($r,4).Font.Bold=$true; $ws.Cells($r,4).HorizontalAlignment=-4108; $ws.Cells($r,4).Interior.Color=$bg
    $ws.Cells($r,5).Value2 = $s.Active; $ws.Cells($r,5).HorizontalAlignment=-4108; $ws.Cells($r,5).Interior.Color=$bg
    $ws.Cells($r,6).Value2 = $s.D07;   $ws.Cells($r,6).HorizontalAlignment=-4108; $ws.Cells($r,6).Interior.Color=$bg
    $ws.Cells($r,7).Value2 = $s.D815;  $ws.Cells($r,7).HorizontalAlignment=-4108; $ws.Cells($r,7).Interior.Color=$bg
    $c16 = $ws.Cells($r,8); $c16.Value2=$s.D1630; $c16.HorizontalAlignment=-4108; $c16.Interior.Color=$bg; if($s.D1630 -gt 0){$c16.Font.Color=$cAmber;$c16.Font.Bold=$true}
    $c30 = $ws.Cells($r,9); $c30.Value2=$s.D30p;  $c30.HorizontalAlignment=-4108; $c30.Interior.Color=$bg; if($s.D30p  -gt 0){$c30.Font.Color=$cRed;$c30.Font.Bold=$true}
    $ws.Cells($r,10).Value2=$s.AvgAge; $ws.Cells($r,10).HorizontalAlignment=-4108; $ws.Cells($r,10).Interior.Color=$bg
    $ws.Cells($r,11).Value2=$s.MaxAge; $ws.Cells($r,11).HorizontalAlignment=-4108; $ws.Cells($r,11).Interior.Color=$bg
    $r++
}

# Team total row
$ws.Rows($r).RowHeight = 18
@(1..11) | ForEach-Object { $ws.Cells($r,$_).Interior.Color = $cNavy; $ws.Cells($r,$_).Font.Color = $cWhite; $ws.Cells($r,$_).Font.Bold = $true }
$ws.Cells($r,1).Value2 = "TEAM TOTAL"; $ws.Cells($r,1).IndentLevel = 1
$ws.Cells($r,2).Value2 = $tIn.ToString("N0"); $ws.Cells($r,2).HorizontalAlignment=-4108
$ws.Cells($r,3).Value2 = $tClosed.ToString("N0"); $ws.Cells($r,3).HorizontalAlignment=-4108
$ws.Cells($r,4).Value2 = "$tRate%"; $ws.Cells($r,4).HorizontalAlignment=-4108
$ws.Cells($r,5).Value2 = $tActive; $ws.Cells($r,5).HorizontalAlignment=-4108
$ws.Cells($r,6).Value2 = $t07;   $ws.Cells($r,6).HorizontalAlignment=-4108
$ws.Cells($r,7).Value2 = $t815;  $ws.Cells($r,7).HorizontalAlignment=-4108
$ws.Cells($r,8).Value2 = $t1630; $ws.Cells($r,8).HorizontalAlignment=-4108
$ws.Cells($r,9).Value2 = $t30p;  $ws.Cells($r,9).HorizontalAlignment=-4108
$ws.Cells($r,10).Value2 = "-"; $ws.Cells($r,10).HorizontalAlignment=-4108
$ws.Cells($r,11).Value2 = "-"; $ws.Cells($r,11).HorizontalAlignment=-4108

$ws.Cells.VerticalAlignment = -4160

# ══════════════════════════════════════════════════════════════
# AGENT SHEETS -- one per agent
# ══════════════════════════════════════════════════════════════
foreach ($s in $agentStats) {
    Write-Host "  Building sheet: $($s.Name)" -ForegroundColor Gray

    $wa = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets($wb.Worksheets.Count))
    # Short tab name
    $tabName = ($s.Name -split ',')[0].Trim()
    if ($tabName.Length -gt 25) { $tabName = $tabName.Substring(0,25) }
    $wa.Name = $tabName

    # Tab color: red if needs attention, green if performing, blue otherwise
    if ($s.D30p -gt 0 -or $s.Rate -lt $TargetRate) { $wa.Tab.Color = $cRed }
    elseif ($s.D30p -eq 0 -and $s.Rate -ge 95)     { $wa.Tab.Color = $cGreen }
    else                                             { $wa.Tab.Color = $cBlue }

    # Column widths
    $wa.Columns(1).ColumnWidth = 22
    $wa.Columns(2).ColumnWidth = 18
    $wa.Columns(3).ColumnWidth = 18
    $wa.Columns(4).ColumnWidth = 18
    $wa.Columns(5).ColumnWidth = 18
    $wa.Columns(6).ColumnWidth = 18

    # Row 1: Agent name header
    $wa.Rows(1).RowHeight = 40
    $r1 = $wa.Range($wa.Cells(1,1), $wa.Cells(1,6))
    $r1.Merge() | Out-Null; $r1.Value2 = $s.Name; $r1.Font.Bold = $true; $r1.Font.Size = 18
    $r1.Font.Color = $cWhite; $r1.Interior.Color = $cNavy; $r1.HorizontalAlignment = -4131; $r1.IndentLevel = 1
    $r1.VerticalAlignment = -4108

    # Row 2: subtitle
    $wa.Rows(2).RowHeight = 18
    $r2 = $wa.Range($wa.Cells(2,1), $wa.Cells(2,6))
    $r2.Merge() | Out-Null; $r2.Value2 = "1:1 Case Review  |  $ReportDate"
    $r2.Font.Size = 10; $r2.Font.Color = (rgb 147 197 253); $r2.Interior.Color = $cNavy; $r2.IndentLevel = 1

    # Row 3: spacer
    $wa.Rows(3).RowHeight = 8

    # Row 4-6: KPI tiles
    $rateColor = if($s.Rate -ge 96){$cGreen}elseif($s.Rate -ge 93){$cBlue}elseif($s.Rate -ge 90){$cAmber}else{$cRed}
    $kpiData = @(
        @{L="YTD Inflow"; V=$s.YTDIn.ToString("N0"); C=$cNavy}
        @{L="YTD Closed"; V=$s.YTDClosed.ToString("N0"); C=$cBlue}
        @{L="YTD Closure Rate"; V="$($s.Rate)%"; C=$rateColor}
        @{L="Active Cases"; V=$s.Active; C=$cNavy}
        @{L="Avg Age"; V=$s.AvgAge; C=if($s.AvgAge -gt 20){$cRed}elseif($s.AvgAge -gt 12){$cAmber}else{$cBlue}}
    )
    $wa.Rows(4).RowHeight = 14
    $wa.Rows(5).RowHeight = 28
    $wa.Rows(6).RowHeight = 5
    for ($i=0; $i -lt $kpiData.Count; $i++) {
        $kd = $kpiData[$i]
        $cl = $wa.Cells(4, $i+1); $cl.Value2 = $kd.L; $cl.Font.Size = 9; $cl.Font.Bold = $true; $cl.Font.Color = $cGray; $cl.Interior.Color = $cLightBg; $cl.HorizontalAlignment = -4131; $cl.IndentLevel = 1
        $cv = $wa.Cells(5, $i+1); $cv.Value2 = $kd.V; $cv.Font.Size = 20; $cv.Font.Bold = $true; $cv.Font.Color = $kd.C; $cv.Interior.Color = $cLightBg; $cv.IndentLevel = 1
        $cc = $wa.Cells(6, $i+1); $cc.Interior.Color = $kd.C
    }

    # Row 7: spacer
    $wa.Rows(7).RowHeight = 8

    # Row 8: Aging label
    $wa.Rows(8).RowHeight = 16
    $rl = $wa.Range($wa.Cells(8,1), $wa.Cells(8,6)); $rl.Merge() | Out-Null
    $rl.Value2 = "ACTIVE CASES BY AGE"; $rl.Font.Bold = $true; $rl.Font.Size = 9; $rl.Font.Color = $cNavy; $rl.IndentLevel = 1

    # Row 9-11: Aging boxes
    $agBoxes = @(
        @{L="0-7 Days";   V=$s.D07;   C=$cNavy; BG=(rgb 239 246 255)}
        @{L="8-15 Days";  V=$s.D815;  C=$cBlue; BG=(rgb 219 234 254)}
        @{L="16-30 Days"; V=$s.D1630; C=if($s.D1630 -gt 0){$cAmber}else{$cGray}; BG=if($s.D1630 -gt 0){$cYellow}else{$cLightBg}}
        @{L=">30 Days";   V=$s.D30p;  C=if($s.D30p -gt 0){$cRed}else{$cGray};   BG=if($s.D30p -gt 0){$cPink}else{$cLightBg}}
    )
    $wa.Rows(9).RowHeight  = 14
    $wa.Rows(10).RowHeight = 26
    $wa.Rows(11).RowHeight = 14
    for ($i=0; $i -lt $agBoxes.Count; $i++) {
        $ab = $agBoxes[$i]
        $cl = $wa.Cells(9,  $i+1); $cl.Value2 = $ab.L;  $cl.Font.Size = 9;  $cl.Font.Bold = $true; $cl.Font.Color = $ab.C; $cl.Interior.Color = $ab.BG; $cl.HorizontalAlignment = -4108
        $cv = $wa.Cells(10, $i+1); $cv.Value2 = $ab.V;  $cv.Font.Size = 22; $cv.Font.Bold = $true; $cv.Font.Color = $ab.C; $cv.Interior.Color = $ab.BG; $cv.HorizontalAlignment = -4108
        $cb = $wa.Cells(11, $i+1); $cb.Interior.Color = $ab.BG
    }

    # Row 12: spacer
    $wa.Rows(12).RowHeight = 8

    # Row 13: Points to discuss label
    $wa.Rows(13).RowHeight = 16
    $rp = $wa.Range($wa.Cells(13,1), $wa.Cells(13,6)); $rp.Merge() | Out-Null
    $rp.Value2 = "POINTS TO DISCUSS"; $rp.Font.Bold = $true; $rp.Font.Size = 9; $rp.Font.Color = $cRed; $rp.Interior.Color = $cPink; $rp.IndentLevel = 1

    $r = 14
    if ($s.D30p -gt 0) {
        $wa.Rows($r).RowHeight = 16
        $rpp = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $rpp.Merge() | Out-Null
        $rpp.Value2 = "  >> $($s.D30p) case(s) older than 30 days -- must be resolved or escalated this week"
        $rpp.Font.Size = 10; $rpp.Font.Bold = $true; $rpp.Font.Color = $cRed; $rpp.Interior.Color = (rgb 255 249 249)
        $r++
    }
    if ($s.D1630 -gt 0) {
        $wa.Rows($r).RowHeight = 16
        $rpp = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $rpp.Merge() | Out-Null
        $rpp.Value2 = "  >> $($s.D1630) case(s) in the 16-30 day bracket -- action required this week"
        $rpp.Font.Size = 10; $rpp.Font.Bold = $true; $rpp.Font.Color = $cAmber; $rpp.Interior.Color = (rgb 255 253 245)
        $r++
    }
    if ($s.Rate -lt $TargetRate) {
        $wa.Rows($r).RowHeight = 16
        $rpp = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $rpp.Merge() | Out-Null
        $rpp.Value2 = "  >> Closure rate is $($s.Rate)% -- below the $TargetRate% team target. Review blockers."
        $rpp.Font.Size = 10; $rpp.Font.Bold = $true; $rpp.Font.Color = $cRed; $rpp.Interior.Color = (rgb 255 249 249)
        $r++
    }
    if ($s.D30p -eq 0 -and $s.D1630 -eq 0 -and $s.Rate -ge $TargetRate) {
        $wa.Rows($r).RowHeight = 16
        $rpp = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $rpp.Merge() | Out-Null
        $rpp.Value2 = "  >> No critical items -- keep up the great work!"
        $rpp.Font.Size = 10; $rpp.Font.Bold = $true; $rpp.Font.Color = $cGreen; $rpp.Interior.Color = $cGreenBg
        $r++
    }

    $r++

    # Coaching notes sections
    $sections = @("Agent Update / Comments","Actions Agreed","Follow-up Required","Coaching Log Entry")
    foreach ($sec in $sections) {
        $wa.Rows($r).RowHeight = 16
        $rh = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $rh.Merge() | Out-Null
        $rh.Value2 = $sec; $rh.Font.Bold = $true; $rh.Font.Size = 10; $rh.Font.Color = $cNavy
        $rh.Interior.Color = $cLightBg; $rh.IndentLevel = 1
        $r++
        # 4 blank input rows
        for ($i=0; $i -lt 4; $i++) {
            $wa.Rows($r).RowHeight = 16
            $ri = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $ri.Merge() | Out-Null
            $ri.Interior.Color = $cWhite
            $ri.Borders.Item(9).LineStyle = 1   # bottom border
            $ri.Borders.Item(9).Color = $cBorder
            $r++
        }
        $wa.Rows($r).RowHeight = 6; $r++
    }

    # Footer
    $wa.Rows($r).RowHeight = 16
    $rf = $wa.Range($wa.Cells($r,1), $wa.Cells($r,6)); $rf.Merge() | Out-Null
    $rf.Value2 = "GTS Management  |  1:1 Review  |  $($s.Name)  |  $ReportDate"
    $rf.Font.Size = 9; $rf.Font.Color = $cGray; $rf.HorizontalAlignment = -4108
    $rf.Borders.Item(9).LineStyle = 1; $rf.Borders.Item(9).Color = $cBorder

    $wa.Cells.VerticalAlignment = -4160
}

# Save and open
$outputFile = "$OutputDir\GTS_1on1_Coaching_Jun17.xlsx"
$wb.SaveAs($outputFile, 51)   # 51 = xlOpenXMLWorkbook
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
