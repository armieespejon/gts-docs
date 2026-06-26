# Add 3-day comparison tab to Active Cases Excel
$xlPath26 = "C:\Users\I535893\Downloads\GTS_Active_Cases_June26.xlsx"

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','COMPLETED_ON','CLOSED_ON','IRT_HOURS')

$deKeywords = [ordered]@{
    'aakriti'='Akanksha Aakriti'; 'medhavi'='Medhavi Maddheshia'
    'mritunjay'='Mritunjay Kumar Singh'; 'mjay'='Mritunjay Kumar Singh'
    'shubham'='Shubham Thakur'; 'shub'='Shubham Thakur'
    'anjali'='Anjali Tripathi'; 'ani'='Anjali Tripathi'
    'swarna'='Swarna Gupta'; 'narendran'='Narendran Raja'; 'naren'='Narendran Raja'
    'kriti'='Kriti Bhalla'; 'ankit'='Ankit Sharan'
    'saksham'='Saksham Pratap Rana'; 'paras'='Paras Arora'
}
$nonTeam = @('Yuan, Michael','Beredo, Randell','Madriaga, Alvin','Nazario, Ma. Charmaine',
             'Espejon, Armie','Santos, Ana Lizelle','Nobela, Monica Frances','Quinto, Marviely')

function Resolve-Name($row) {
    if ($row.PROCESSOR_ID -eq 'SAP_GERMANY_USER') {
        $misc = $row.EXT_MISCINFO.Trim().ToLower()
        foreach ($k in $deKeywords.Keys) { if ($misc -match $k) { return $deKeywords[$k] } }
        return 'Germany (Unassigned)'
    }
    $n = $row.PROCESSOR_NAME.Trim()
    if ([string]::IsNullOrWhiteSpace($n) -or $n -eq 'None') { return 'Unassigned' }
    if ($n -in $nonTeam) { return 'Others' }
    return $n
}

function Get-ActiveCounts($csvPath) {
    $data = Import-Csv $csvPath -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }
    $active = $data | Where-Object { $_.STATUS -notin @('Closed','Completed','Customer Action') }
    $counts = @{}
    foreach ($row in $active) {
        $n = Resolve-Name $row
        if (-not $counts.ContainsKey($n)) { $counts[$n] = 0 }
        $counts[$n]++
    }
    return $counts
}

$jun24 = Get-ActiveCounts "$env:TEMP\gts_cases_june24.csv"
$jun25 = Get-ActiveCounts "$env:TEMP\gts_cases_june25.csv"
$jun26 = Get-ActiveCounts "$env:TEMP\gts_cases_june26.csv"

$allAgents = ($jun24.Keys + $jun25.Keys + $jun26.Keys) | Sort-Object -Unique

# Open Excel and add tab
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlPath26)

# Remove existing tab if present
foreach ($sh in @($wb.Sheets | Where-Object { $_.Name -eq 'Jun24-25-26 Comparison' })) { $sh.Delete() }

$ws = $wb.Sheets.Add()
$ws.Name = 'Jun24-25-26 Comparison'
$wb.Sheets('Jun24-25-26 Comparison').Move($wb.Sheets($wb.Sheets.Count)) | Out-Null

# Headers
$hdrs = @('Agent','Jun 24','Jun 25','Jun 26','Change','Trend')
for ($c = 1; $c -le $hdrs.Count; $c++) {
    $ws.Cells(1,$c).Value2 = $hdrs[$c-1]
    $ws.Cells(1,$c).Font.Bold = $true
    $ws.Cells(1,$c).Interior.Color = 0x5F3A1E
    $ws.Cells(1,$c).Font.Color = 0xFFFFFF
}

$r = 2
$tot24 = 0; $tot25 = 0; $tot26 = 0

foreach ($a in $allAgents) {
    $d24 = if ($jun24.ContainsKey($a)) { $jun24[$a] } else { 0 }
    $d25 = if ($jun25.ContainsKey($a)) { $jun25[$a] } else { 0 }
    $d26 = if ($jun26.ContainsKey($a)) { $jun26[$a] } else { 0 }
    $diff = $d26 - $d24
    $diffStr = if ($diff -gt 0) { "+$diff" } elseif ($diff -lt 0) { "$diff" } else { "0" }
    $trend = if ($diff -lt 0) { 'Improved' } elseif ($diff -gt 0) { 'Increased' } else { 'No change' }

    $ws.Cells($r,1).Value2 = $a
    $ws.Cells($r,2).Value2 = "$d24"
    $ws.Cells($r,3).Value2 = "$d25"
    $ws.Cells($r,4).Value2 = "$d26"
    $ws.Cells($r,5).Value2 = $diffStr
    $ws.Cells($r,6).Value2 = $trend

    if ($diff -lt 0) {
        $ws.Cells($r,5).Interior.Color = 0xC6EFCE; $ws.Cells($r,5).Font.Color = 0x276221
        $ws.Cells($r,6).Interior.Color = 0xC6EFCE; $ws.Cells($r,6).Font.Color = 0x276221
        $ws.Cells($r,6).Font.Bold = $true
    } elseif ($diff -gt 0) {
        $ws.Cells($r,5).Interior.Color = 0xFFCCCC; $ws.Cells($r,5).Font.Color = 0x991b1b
    }

    $tot24 += $d24; $tot25 += $d25; $tot26 += $d26
    $r++
}

# Total row
$totDiff = $tot26 - $tot24
$totDiffStr = if ($totDiff -gt 0) { "+$totDiff" } else { "$totDiff" }
$ws.Cells($r,1).Value2 = 'TOTAL'
$ws.Cells($r,1).Font.Bold = $true
$ws.Cells($r,2).Value2 = "$tot24"
$ws.Cells($r,3).Value2 = "$tot25"
$ws.Cells($r,4).Value2 = "$tot26"
$ws.Cells($r,5).Value2 = $totDiffStr
$ws.Cells($r,6).Value2 = ''
for ($c = 1; $c -le 6; $c++) {
    $ws.Cells($r,$c).Font.Bold = $true
    $ws.Cells($r,$c).Interior.Color = 0x5F3A1E
    $ws.Cells($r,$c).Font.Color = 0xFFFFFF
}

$ws.Columns.AutoFit() | Out-Null
$wb.Save()
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Done! 3-day comparison tab added."
Start-Process $xlPath26
