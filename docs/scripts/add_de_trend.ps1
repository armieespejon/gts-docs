# Add New DE Daily Trend tab to Active Cases Excel
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

function Resolve-Name($row) {
    if ($row.PROCESSOR_ID -eq 'SAP_GERMANY_USER') {
        $misc = $row.EXT_MISCINFO.Trim().ToLower()
        foreach ($k in $deKeywords.Keys) { if ($misc -match $k) { return $deKeywords[$k] } }
        return $null
    }
    return $null
}

$data = Import-Csv "$env:TEMP\gts_cases_june26.csv" -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }
$newDE = @('Saksham Pratap Rana','Ankit Sharan','Akanksha Aakriti','Medhavi Maddheshia','Paras Arora')
$days  = 18..26

# Build data
$agentData = @{}
foreach ($agent in $newDE) {
    $agentData[$agent] = @{}
    $agentCases = $data | Where-Object { (Resolve-Name $_) -eq $agent }
    foreach ($d in $days) {
        $cnt = ($agentCases | Where-Object {
            try { [datetime]::Parse($_.CREATED_ON_DATE).Day -eq $d } catch { $false }
        }).Count
        $agentData[$agent][$d] = $cnt
    }
}

# Open Excel
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlPath26)

# Remove existing tab if present
foreach ($sh in @($wb.Sheets | Where-Object { $_.Name -eq 'New DE Daily Trend' })) { $sh.Delete() }

$ws = $wb.Sheets.Add()
$ws.Name = 'New DE Daily Trend'
$wb.Sheets('New DE Daily Trend').Move($wb.Sheets($wb.Sheets.Count)) | Out-Null

# Headers
$ws.Cells(1,1).Value2 = 'Agent'
$ws.Cells(1,1).Font.Bold = $true
$ws.Cells(1,1).Interior.Color = 0x5F3A1E
$ws.Cells(1,1).Font.Color = 0xFFFFFF
$c = 2
foreach ($d in $days) {
    $ws.Cells(1,$c).Value2 = "Jun $d"
    $ws.Cells(1,$c).Font.Bold = $true
    $ws.Cells(1,$c).Interior.Color = 0x5F3A1E
    $ws.Cells(1,$c).Font.Color = 0xFFFFFF
    $c++
}
$ws.Cells(1,$c).Value2 = 'Running Total'
$ws.Cells(1,$c).Font.Bold = $true
$ws.Cells(1,$c).Interior.Color = 0x14532D
$ws.Cells(1,$c).Font.Color = 0xFFFFFF

# Daily totals row data
$dayTotals = @{}
foreach ($d in $days) { $dayTotals[$d] = 0 }
$grandTotal = 0

# Data rows
$r = 2
foreach ($agent in $newDE) {
    $ws.Cells($r,1).Value2 = $agent
    $c = 2
    $runTotal = 0
    foreach ($d in $days) {
        $v = $agentData[$agent][$d]
        if ($v -gt 0) {
            $ws.Cells($r,$c).Value2 = "$v"
            $ws.Cells($r,$c).Interior.Color = 0xD1FAE5
        } else {
            $ws.Cells($r,$c).Value2 = '--'
            $ws.Cells($r,$c).Font.Color = 0xCCCCCC
        }
        $runTotal += $v
        $dayTotals[$d] += $v
        $c++
    }
    $ws.Cells($r,$c).Value2 = "$runTotal"
    $ws.Cells($r,$c).Font.Bold = $true
    $ws.Cells($r,$c).Interior.Color = 0xD1FAE5
    $grandTotal += $runTotal
    $r++
}

# Team total row
$ws.Cells($r,1).Value2 = 'TEAM TOTAL'
$ws.Cells($r,1).Font.Bold = $true
$ws.Cells($r,1).Interior.Color = 0x14532D
$ws.Cells($r,1).Font.Color = 0xFFFFFF
$c = 2
foreach ($d in $days) {
    $ws.Cells($r,$c).Value2 = "$($dayTotals[$d])"
    $ws.Cells($r,$c).Font.Bold = $true
    $ws.Cells($r,$c).Interior.Color = 0x14532D
    $ws.Cells($r,$c).Font.Color = 0xFFFFFF
    $c++
}
$ws.Cells($r,$c).Value2 = "$grandTotal"
$ws.Cells($r,$c).Font.Bold = $true
$ws.Cells($r,$c).Interior.Color = 0x14532D
$ws.Cells($r,$c).Font.Color = 0xFFFFFF

$ws.Columns.AutoFit() | Out-Null
$wb.Save()
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Done! New DE Daily Trend tab added."
Start-Process $xlPath26
