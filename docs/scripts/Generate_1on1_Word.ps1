# ============================================================
# GTS 1:1 Coaching Review -- Simple Word Generator
# ============================================================

# -- CONFIG ---------------------------------------------------
$CasesXlsx  = "C:\Users\I535893\Downloads\GTS ESM Cases_June17.xlsx"
$OutputFile = "C:\Users\I535893\Downloads\GTS_1on1_Coaching_Jun17.docx"
$ReportDate = "June 17, 2026"
$Today      = [datetime]'2026-06-17'
$TargetRate = 90
# -- END CONFIG -----------------------------------------------

$CasesCsv = "$env:TEMP\gts_1on1.csv"

# Export xlsx to csv
Write-Host "Exporting data..." -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($CasesXlsx)
$wb.SaveAs($CasesCsv, 6)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE','EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13','L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY','CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')
$cases = Import-Csv $CasesCsv -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

Write-Host "Computing stats..." -ForegroundColor Cyan

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

$stats = @()
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
    $avg = if ($ages.Count -gt 0) { [math]::Round(($ages | Measure-Object -Average).Average,1) } else { 0 }
    $max = if ($ages.Count -gt 0) { ($ages | Measure-Object -Maximum).Maximum } else { 0 }
    $stats += [PSCustomObject]@{ Name=$name; In=$a.YTDIn; Closed=$a.YTDClosed; Rate=$rate; Active=$a.Active.Count; D07=$d07; D815=$d815; D1630=$d1630; D30p=$d30p; Avg=$avg; Max=$max }
}

$tIn     = ($stats | Measure-Object In     -Sum).Sum
$tClosed = ($stats | Measure-Object Closed -Sum).Sum
$tActive = ($stats | Measure-Object Active -Sum).Sum
$tRate   = [math]::Round($tClosed/$tIn*100)
$t07     = ($stats | Measure-Object D07    -Sum).Sum
$t815    = ($stats | Measure-Object D815   -Sum).Sum
$t1630   = ($stats | Measure-Object D1630  -Sum).Sum
$t30p    = ($stats | Measure-Object D30p   -Sum).Sum

Write-Host "Building Word doc..." -ForegroundColor Cyan

$wd  = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$sel = $wd.Selection

# Page setup
$doc.PageSetup.TopMargin    = $wd.CentimetersToPoints(1.5)
$doc.PageSetup.BottomMargin = $wd.CentimetersToPoints(1.5)
$doc.PageSetup.LeftMargin   = $wd.CentimetersToPoints(1.8)
$doc.PageSetup.RightMargin  = $wd.CentimetersToPoints(1.8)

# Color constants (Word BGR long)
$WHITE  = [int]16777215
$BLACK  = [int]0
$NAVY   = [int]6245150    # 30,58,95
$BLUE   = [int]14163997   # 29,78,216
$GREEN  = [int]5742596    # 4,120,87
$RED    = [int]1843129    # 185,28,28
$AMBER  = [int]25780      # 180,100,0
$GRAY   = [int]9145964    # 100,116,139

function Type-Heading($text, $size=14, $col=$NAVY) {
    $sel.TypeText($text)
    $p = $sel.Paragraphs.Last
    $p.Range.Font.Size  = $size
    $p.Range.Font.Bold  = $true
    $p.Range.Font.Color = $col
    $p.SpaceAfter  = 4
    $p.SpaceBefore = 8
    $sel.TypeParagraph()
}

function Type-Sub($text, $size=10, $col=$GRAY) {
    $sel.TypeText($text)
    $p = $sel.Paragraphs.Last
    $p.Range.Font.Size  = $size
    $p.Range.Font.Bold  = $false
    $p.Range.Font.Color = $col
    $p.SpaceAfter = 6
    $sel.TypeParagraph()
}

function Type-Point($text, $col=$RED) {
    $sel.TypeText("  * " + $text)
    $p = $sel.Paragraphs.Last
    $p.Range.Font.Size  = 10
    $p.Range.Font.Bold  = $true
    $p.Range.Font.Color = $col
    $p.SpaceAfter = 3
    $sel.TypeParagraph()
}

function Add-KPI-Row($tbl, $row, $vals, $colors) {
    for ($c = 1; $c -le $vals.Count; $c++) {
        $cell = $tbl.Cell($row, $c)
        $cell.Range.Text = $vals[$c-1]
        $cell.Range.Font.Bold  = $true
        $cell.Range.Font.Size  = if ($row -eq 2) { 20 } else { 9 }
        $cell.Range.Font.Color = if ($row -eq 2 -and $colors -and $colors[$c-1]) { $colors[$c-1] } else { $GRAY }
        $cell.Range.ParagraphFormat.Alignment = 1  # center
        $cell.Shading.BackgroundPatternColor  = [int]16316664  # light bg
    }
}

function Fill-Table-Header($tbl, $hdrs) {
    for ($c = 1; $c -le $hdrs.Count; $c++) {
        $cell = $tbl.Cell(1, $c)
        $cell.Range.Text = $hdrs[$c-1]
        $cell.Range.Font.Bold  = $true
        $cell.Range.Font.Size  = 10
        $cell.Range.Font.Color = $WHITE
        $cell.Shading.BackgroundPatternColor = $NAVY
        $cell.Range.ParagraphFormat.Alignment = if ($c -eq 1) { 0 } else { 1 }
    }
}

function Fill-Table-Row($tbl, $row, $vals, $bg) {
    for ($c = 1; $c -le $vals.Count; $c++) {
        $cell = $tbl.Cell($row, $c)
        $cell.Range.Text = "$($vals[$c-1])"
        $cell.Range.Font.Size = 10
        $cell.Shading.BackgroundPatternColor = $bg
        $cell.Range.ParagraphFormat.Alignment = if ($c -eq 1) { 0 } else { 1 }
    }
}

function Rate-Color($r) { if ($r -ge 96) { return $GREEN } elseif ($r -ge 93) { return $BLUE } elseif ($r -ge 90) { return $AMBER } else { return $RED } }

# ── PAGE 1: MANAGER SUMMARY ───────────────────────────────────

Type-Heading "GTS Team — Manager Summary & 1:1 Review" 18 $NAVY
Type-Sub "As of $ReportDate  |  GTS Management  |  Confidential" 10 $GRAY

# Team KPI table
Type-Heading "Team Performance Snapshot" 12 $BLUE
$t1 = $doc.Tables.Add($sel.Range, 2, 4)
$t1.Style = "Table Grid"
Fill-Table-Header $t1 @("YTD Inflow","YTD Closures","YTD Closure Rate","Active Cases")
$rateColor = Rate-Color $tRate
Fill-Table-Row $t1 2 @($tIn.ToString("N0"), $tClosed.ToString("N0"), "$tRate%", "$tActive") [int]16316664
$t1.Cell(2,1).Range.Font.Color = $NAVY;  $t1.Cell(2,1).Range.Font.Size = 20; $t1.Cell(2,1).Range.Font.Bold = $true
$t1.Cell(2,2).Range.Font.Color = $BLUE;  $t1.Cell(2,2).Range.Font.Size = 20; $t1.Cell(2,2).Range.Font.Bold = $true
$t1.Cell(2,3).Range.Font.Color = $rateColor; $t1.Cell(2,3).Range.Font.Size = 20; $t1.Cell(2,3).Range.Font.Bold = $true
$t1.Cell(2,4).Range.Font.Color = $AMBER; $t1.Cell(2,4).Range.Font.Size = 20; $t1.Cell(2,4).Range.Font.Bold = $true
for ($c=1; $c -le 4; $c++) { $t1.Columns($c).Width = $wd.CentimetersToPoints(4.0) }
$sel.MoveDown(5,1) | Out-Null; $sel.TypeParagraph()

# Aging table
Type-Heading "Active Cases by Age" 11 $GRAY
$t2 = $doc.Tables.Add($sel.Range, 2, 4)
$t2.Style = "Table Grid"
Fill-Table-Header $t2 @("0-7 Days","8-15 Days","16-30 Days",">30 Days")
Fill-Table-Row $t2 2 @("$t07","$t815","$t1630","$t30p") [int]16316664
$t2.Cell(2,1).Range.Font.Color = $NAVY;  $t2.Cell(2,1).Range.Font.Size = 22; $t2.Cell(2,1).Range.Font.Bold = $true
$t2.Cell(2,2).Range.Font.Color = $BLUE;  $t2.Cell(2,2).Range.Font.Size = 22; $t2.Cell(2,2).Range.Font.Bold = $true
$t2.Cell(2,3).Range.Font.Color = $AMBER; $t2.Cell(2,3).Range.Font.Size = 22; $t2.Cell(2,3).Range.Font.Bold = $true
$t2.Cell(2,4).Range.Font.Color = $RED;   $t2.Cell(2,4).Range.Font.Size = 22; $t2.Cell(2,4).Range.Font.Bold = $true
for ($c=1; $c -le 4; $c++) { $t2.Columns($c).Width = $wd.CentimetersToPoints(4.0) }
$sel.MoveDown(5,1) | Out-Null; $sel.TypeParagraph()

# Needs Attention
$needsAttn = $stats | Where-Object { $_.D30p -gt 0 -or $_.Rate -lt $TargetRate } | Sort-Object { $_.D30p*100 + $_.D1630 } -Descending
$naLabel = "Needs Attention (" + $needsAttn.Count + " agents)"
Type-Heading $naLabel 11 $RED
$tNA = $doc.Tables.Add($sel.Range, ($needsAttn.Count+1), 5)
$tNA.Style = "Table Grid"
Fill-Table-Header $tNA @("Agent","Active",">30d","16-30d","Closure Rate")
$ri = 2
foreach ($s in $needsAttn) {
    $bg = if ($ri % 2 -eq 0) { [int]16119283 } else { $WHITE }
    Fill-Table-Row $tNA $ri @($s.Name,"$($s.Active)","$($s.D30p)","$($s.D1630)","$($s.Rate)%") $bg
    $tNA.Cell($ri,1).Range.Font.Bold = $true
    $tNA.Cell($ri,5).Range.Font.Color = Rate-Color $s.Rate
    $tNA.Cell($ri,5).Range.Font.Bold = $true
    if ($s.D30p  -gt 0) { $tNA.Cell($ri,3).Range.Font.Color=$RED;   $tNA.Cell($ri,3).Range.Font.Bold=$true }
    if ($s.D1630 -gt 0) { $tNA.Cell($ri,4).Range.Font.Color=$AMBER; $tNA.Cell($ri,4).Range.Font.Bold=$true }
    $ri++
}
$tNA.Columns(1).Width = $wd.CentimetersToPoints(5.5)
for ($c=2; $c -le 5; $c++) { $tNA.Columns($c).Width = $wd.CentimetersToPoints(2.0) }
$sel.MoveDown(5,1) | Out-Null; $sel.TypeParagraph()

# Performing Well
$performing = $stats | Where-Object { $_.D30p -eq 0 -and $_.Rate -ge 95 } | Sort-Object Rate -Descending
$pwLabel = "Performing Well (" + $performing.Count + " agents)"
Type-Heading $pwLabel 11 $GREEN
$tPW = $doc.Tables.Add($sel.Range, ($performing.Count+1), 4)
$tPW.Style = "Table Grid"
Fill-Table-Header $tPW @("Agent","Active","Avg Age","Closure Rate")
$ri = 2
foreach ($s in $performing) {
    $bg = if ($ri % 2 -eq 0) { [int]14680064 } else { $WHITE }
    Fill-Table-Row $tPW $ri @($s.Name,"$($s.Active)","$($s.Avg)","$($s.Rate)%") $bg
    $tPW.Cell($ri,1).Range.Font.Color = $GREEN; $tPW.Cell($ri,1).Range.Font.Bold = $true
    $tPW.Cell($ri,4).Range.Font.Color = $GREEN; $tPW.Cell($ri,4).Range.Font.Bold = $true
    $ri++
}
$tPW.Columns(1).Width = $wd.CentimetersToPoints(5.5)
for ($c=2; $c -le 4; $c++) { $tPW.Columns($c).Width = $wd.CentimetersToPoints(2.5) }
$sel.MoveDown(5,1) | Out-Null; $sel.TypeParagraph()

# Full team - new page
$sel.InsertBreak(7)
Type-Heading "Full Team Overview" 12 $NAVY
$tFT = $doc.Tables.Add($sel.Range, ($stats.Count+2), 11)
$tFT.Style = "Table Grid"
Fill-Table-Header $tFT @("Agent","YTD In","Closed","Rate","Active","0-7d","8-15d","16-30d",">30d","Avg","Max")
$ri = 2
foreach ($s in $stats) {
    $bg = if ($ri % 2 -eq 0) { [int]16316664 } else { $WHITE }
    Fill-Table-Row $tFT $ri @($s.Name,"$($s.In)","$($s.Closed)","$($s.Rate)%","$($s.Active)","$($s.D07)","$($s.D815)","$($s.D1630)","$($s.D30p)","$($s.Avg)","$($s.Max)") $bg
    $tFT.Cell($ri,4).Range.Font.Color = Rate-Color $s.Rate; $tFT.Cell($ri,4).Range.Font.Bold = $true
    if ($s.D1630 -gt 0) { $tFT.Cell($ri,8).Range.Font.Color=$AMBER; $tFT.Cell($ri,8).Range.Font.Bold=$true }
    if ($s.D30p  -gt 0) { $tFT.Cell($ri,9).Range.Font.Color=$RED;   $tFT.Cell($ri,9).Range.Font.Bold=$true }
    $ri++
}
Fill-Table-Row $tFT $ri @("TEAM TOTAL",$tIn.ToString("N0"),$tClosed.ToString("N0"),"$tRate%","$tActive","$t07","$t815","$t1630","$t30p","-","-") $NAVY
for ($c=1; $c -le 11; $c++) { $tFT.Cell($ri,$c).Range.Font.Color=$WHITE; $tFT.Cell($ri,$c).Range.Font.Bold=$true }
$tFT.Columns(1).Width = $wd.CentimetersToPoints(4.5)
for ($c=2; $c -le 11; $c++) { $tFT.Columns($c).Width = $wd.CentimetersToPoints(1.5) }
$sel.MoveDown(5,1) | Out-Null

# ── INDIVIDUAL 1:1 SHEETS ─────────────────────────────────────
foreach ($s in $stats) {
    Write-Host "  Sheet: $($s.Name)" -ForegroundColor Gray
    $sel.InsertBreak(7)

    # Name header
    Type-Heading $s.Name 18 $BLUE
    Type-Sub "1:1 Case Review  |  $ReportDate" 10 $GRAY

    # KPI table
    $tk = $doc.Tables.Add($sel.Range, 2, 5)
    $tk.Style = "Table Grid"
    Fill-Table-Header $tk @("YTD Inflow","YTD Closed","YTD Closure Rate","Active Cases","Avg Age")
    Fill-Table-Row $tk 2 @($s.In.ToString("N0"),$s.Closed.ToString("N0"),"$($s.Rate)%","$($s.Active)","$($s.Avg)") [int]16316664
    $tk.Cell(2,1).Range.Font.Color = $NAVY;  $tk.Cell(2,1).Range.Font.Size=18; $tk.Cell(2,1).Range.Font.Bold=$true
    $tk.Cell(2,2).Range.Font.Color = $BLUE;  $tk.Cell(2,2).Range.Font.Size=18; $tk.Cell(2,2).Range.Font.Bold=$true
    $tk.Cell(2,3).Range.Font.Color = Rate-Color $s.Rate; $tk.Cell(2,3).Range.Font.Size=18; $tk.Cell(2,3).Range.Font.Bold=$true
    $tk.Cell(2,4).Range.Font.Color = $NAVY;  $tk.Cell(2,4).Range.Font.Size=18; $tk.Cell(2,4).Range.Font.Bold=$true
    $avgCol = if($s.Avg -gt 20){$RED}elseif($s.Avg -gt 12){$AMBER}else{$BLUE}
    $tk.Cell(2,5).Range.Font.Color = $avgCol; $tk.Cell(2,5).Range.Font.Size=18; $tk.Cell(2,5).Range.Font.Bold=$true
    for ($c=1; $c -le 5; $c++) { $tk.Columns($c).Width = $wd.CentimetersToPoints(3.2) }
    $sel.MoveDown(5,1) | Out-Null; $sel.TypeParagraph()

    # Aging table
    Type-Sub "Active Cases by Age" 10 $GRAY
    $ta = $doc.Tables.Add($sel.Range, 2, 4)
    $ta.Style = "Table Grid"
    Fill-Table-Header $ta @("0-7 Days","8-15 Days","16-30 Days",">30 Days")
    Fill-Table-Row $ta 2 @("$($s.D07)","$($s.D815)","$($s.D1630)","$($s.D30p)") [int]16316664
    $ta.Cell(2,1).Range.Font.Color=$NAVY;  $ta.Cell(2,1).Range.Font.Size=22; $ta.Cell(2,1).Range.Font.Bold=$true
    $ta.Cell(2,2).Range.Font.Color=$BLUE;  $ta.Cell(2,2).Range.Font.Size=22; $ta.Cell(2,2).Range.Font.Bold=$true
    $c16col = if($s.D1630 -gt 0){$AMBER}else{$GRAY}
    $c30col = if($s.D30p  -gt 0){$RED}  else{$GRAY}
    $ta.Cell(2,3).Range.Font.Color=$c16col; $ta.Cell(2,3).Range.Font.Size=22; $ta.Cell(2,3).Range.Font.Bold=$true
    $ta.Cell(2,4).Range.Font.Color=$c30col; $ta.Cell(2,4).Range.Font.Size=22; $ta.Cell(2,4).Range.Font.Bold=$true
    for ($c=1; $c -le 4; $c++) { $ta.Columns($c).Width = $wd.CentimetersToPoints(4.0) }
    $sel.MoveDown(5,1) | Out-Null; $sel.TypeParagraph()

    # Points to discuss
    $sel.TypeText("Points to Discuss")
    $sel.Paragraphs.Last.Range.Font.Bold=$true; $sel.Paragraphs.Last.Range.Font.Size=10
    $sel.Paragraphs.Last.Range.Font.Color=$RED; $sel.Paragraphs.Last.SpaceAfter=4; $sel.Paragraphs.Last.SpaceBefore=8
    $sel.TypeParagraph()

    if ($s.D30p -gt 0)              { Type-Point "$($s.D30p) case(s) older than 30 days -- must be resolved or escalated." $RED }
    if ($s.D1630 -gt 0)             { Type-Point "$($s.D1630) case(s) in the 16-30 day bracket -- action required this week." $AMBER }
    if ($s.Rate -lt $TargetRate)    { Type-Point "Closure rate is $($s.Rate)% -- below the $TargetRate% target. Review blockers." $RED }
    if ($s.D30p -eq 0 -and $s.D1630 -eq 0 -and $s.Rate -ge $TargetRate) { Type-Point "No critical items -- keep up the great work!" $GREEN }

    # Notes sections
    $sel.TypeParagraph()
    foreach ($section in @("1:1 Notes","Agent Update / Comments:","Actions Agreed:","Follow-up Required:","Coaching Log Entry:")) {
        $sel.TypeText($section)
        $sel.Paragraphs.Last.Range.Font.Bold=$true; $sel.Paragraphs.Last.Range.Font.Size=10
        $sel.Paragraphs.Last.Range.Font.Color=$NAVY; $sel.Paragraphs.Last.SpaceBefore=8; $sel.Paragraphs.Last.SpaceAfter=2
        $sel.TypeParagraph()
        for ($i=1; $i -le 3; $i++) {
            $sel.TypeText(" ")
            $sel.Paragraphs.Last.Range.Font.Size=11
            $brd = $sel.Paragraphs.Last.Borders.Item(3)
            $brd.LineStyle=1; $brd.Weight=1; $brd.Color=[int]13882323
            $sel.Paragraphs.Last.SpaceAfter=10
            $sel.TypeParagraph()
        }
    }

    # Footer
    $sel.TypeText("GTS Management  |  1:1 Review  |  $($s.Name)  |  $ReportDate")
    $sel.Paragraphs.Last.Range.Font.Size=9; $sel.Paragraphs.Last.Range.Font.Color=$GRAY
    $sel.Paragraphs.Last.Alignment=1; $sel.Paragraphs.Last.SpaceBefore=10
    $sel.TypeParagraph()
}

# Save
Write-Host "Saving..." -ForegroundColor Cyan
$doc.SaveAs([ref]$OutputFile, [ref]16)
$doc.Close($false)
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd) | Out-Null

Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $OutputFile" -ForegroundColor Yellow
Start-Process $OutputFile
