# Export Active Cases to Excel
$csvPath  = "$env:TEMP\gts_cases_inflow.csv"
$xlPath   = "C:\Users\I535893\Downloads\GTS_Active_Cases_June26.xlsx"
$today    = [datetime]::Today

$cHdr = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6','EXT_COURSECODE',
          'EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS','SOURCE','SERVICETEAMNAME','COL13',
          'L1_CATEGORY','L2_CATEGORY','ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY',
          'CONTENT_TEAM','PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','COMPLETED_ON','CLOSED_ON','IRT_HOURS')
$data = Import-Csv $csvPath -Header $cHdr | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }

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

$active = $data | Where-Object { $_.STATUS -notin @('Closed','Completed','Customer Action') }

$rows = $active | ForEach-Object {
    $age = try { [int]($today - [datetime]::Parse($_.CREATED_ON_DATE)).Days } catch { 0 }
    [PSCustomObject]@{
        Agent   = Resolve-Name $_
        CaseID  = $_.DISPLAYID
        Status  = $_.STATUS
        Created = $_.CREATED_ON_DATE
        Age     = $age
        Bucket  = if($age -ge 31){'30+ days'}elseif($age -ge 16){'16-30 days'}elseif($age -ge 8){'8-15 days'}else{'0-7 days'}
        Course  = $_.EXT_COURSECODE
        L1      = $_.L1_CATEGORY
        Subject = $_.SUBJECT
    }
} | Sort-Object Agent, { -$_.Age }

# Build summary
$agentList = $rows | Select-Object -ExpandProperty Agent | Sort-Object -Unique
$summary = $agentList | ForEach-Object {
    $a = $_
    $aRows = $rows | Where-Object { $_.Agent -eq $a }
    [PSCustomObject]@{
        Agent      = $a
        Total      = $aRows.Count
        D0_7       = ($aRows | Where-Object Bucket -eq '0-7 days').Count
        D8_15      = ($aRows | Where-Object Bucket -eq '8-15 days').Count
        D16_30     = ($aRows | Where-Object Bucket -eq '16-30 days').Count
        D30plus    = ($aRows | Where-Object Bucket -eq '30+ days').Count
    }
}

# Write Excel
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

# ---- Summary Sheet ----
$ws1 = $wb.Sheets(1); $ws1.Name = 'Summary'
$hdr = @('Agent','Total','0-7 days','8-15 days','16-30 days','30+ days')
for ($c = 1; $c -le $hdr.Count; $c++) {
    $cell = $ws1.Cells(1,$c)
    $cell.Value2 = $hdr[$c-1]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 0x5F3A1E
    $cell.Font.Color = 0xFFFFFF
}
$r = 2
foreach ($s in $summary) {
    $ws1.Cells($r,1).Value2 = $s.Agent
    $ws1.Cells($r,2).Value2 = "$($s.Total)";   $ws1.Cells($r,2).Font.Bold = $true
    $ws1.Cells($r,3).Value2 = "$($s.D0_7)"
    $ws1.Cells($r,4).Value2 = "$($s.D8_15)";   if($s.D8_15  -gt 0){ $ws1.Cells($r,4).Interior.Color = 0xCCFFFF }
    $ws1.Cells($r,5).Value2 = "$($s.D16_30)";  if($s.D16_30 -gt 0){ $ws1.Cells($r,5).Interior.Color = 0x99CCFF }
    $ws1.Cells($r,6).Value2 = "$($s.D30plus)"; if($s.D30plus -gt 0){ $ws1.Cells($r,6).Interior.Color = 0x9999FF; $ws1.Cells($r,6).Font.Bold = $true }
    $r++
}
# Total row
$ws1.Cells($r,1).Value2 = 'TOTAL'; $ws1.Cells($r,1).Font.Bold = $true
$ws1.Cells($r,1).Interior.Color = 0x5F3A1E; $ws1.Cells($r,1).Font.Color = 0xFFFFFF
foreach ($c in 2..6) {
    $prop = @('Total','D0_7','D8_15','D16_30','D30plus')[$c-2]
    $sum = ($summary | Measure-Object -Property $prop -Sum).Sum
    $ws1.Cells($r,$c).Value2 = "$sum"
    $ws1.Cells($r,$c).Font.Bold = $true
    $ws1.Cells($r,$c).Interior.Color = 0x5F3A1E
    $ws1.Cells($r,$c).Font.Color = 0xFFFFFF
}
$ws1.Columns.AutoFit() | Out-Null

# ---- Active Cases Sheet ----
$ws2 = $wb.Sheets.Add()
$ws2.Name = 'Active Cases'
$wb.Sheets('Active Cases').Move($wb.Sheets(1)) | Out-Null
$hdrs = @('Agent','Case ID','Status','Created Date','Age (Days)','Bucket','Course','L1 Category','Subject')
for ($c = 1; $c -le $hdrs.Count; $c++) {
    $cell = $ws2.Cells(1,$c)
    $cell.Value2 = $hdrs[$c-1]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 0x5F3A1E
    $cell.Font.Color = 0xFFFFFF
}
$r = 2
foreach ($row in $rows) {
    $ws2.Cells($r,1).Value2 = $row.Agent
    $ws2.Cells($r,2).Value2 = $row.CaseID
    $ws2.Cells($r,3).Value2 = $row.Status
    $ws2.Cells($r,4).Value2 = $row.Created
    $ws2.Cells($r,5).Value2 = "$($row.Age)"
    $ws2.Cells($r,6).Value2 = $row.Bucket
    $ws2.Cells($r,7).Value2 = $row.Course
    $ws2.Cells($r,8).Value2 = $row.L1
    $ws2.Cells($r,9).Value2 = $row.Subject
    $bgColor = switch($row.Bucket) {
        '30+ days'   { 0x9999FF }
        '16-30 days' { 0x99CCFF }
        '8-15 days'  { 0xCCFFFF }
        default      { 0xFFFFFF }
    }
    $ws2.Cells($r,5).Interior.Color = $bgColor
    $ws2.Cells($r,6).Interior.Color = $bgColor
    $r++
}
$ws2.Columns.AutoFit() | Out-Null

$wb.SaveAs($xlPath)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Done! $($rows.Count) active cases exported to $xlPath"
Start-Process $xlPath
