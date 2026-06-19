# ============================================================
# GTS Winners Circle — Auto-Generator
# Usage: .\Generate_Winners_Circle.ps1
#
# 1. Drop the new weekly xlsx into Downloads (any name is fine,
#    update $XlsxPath below to match).
# 2. Set the week label and date range at the top.
# 3. Run. HTML is written to the same folder.
# ============================================================

# ── CONFIG ──────────────────────────────────────────────────
$XlsxPath    = "C:\Users\I535893\Downloads\GTS ESM Cases_June16.xlsx"
$WeekLabel   = "W25"
$WeekRange   = "Jun 15–21, 2026"
$NextBulletin = "Monday, June 29"
$RunDate     = "June 16, 2026"   # "As of" date shown in the table header
$TipOfWeek   = "Start every shift by sorting your queue oldest first. Your aged cases are your first priority — not the newest ones. A clean backlog is a happy customer. 🏆"
# ── END CONFIG ───────────────────────────────────────────────

$CsvPath = [System.IO.Path]::ChangeExtension($XlsxPath, ".tmp.csv")

# ── Step 1: Export xlsx → CSV via Excel COM ─────────────────
Write-Host "Exporting Excel to CSV..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($XlsxPath)
$wb.SaveAs($CsvPath, 6)
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "  Done." -ForegroundColor Green

# ── Step 2: Load data ────────────────────────────────────────
$headers = @('DISPLAYID','SUBJECT','YEAR','MONTH','CREATED_ON_DATE','COL6',
             'EXT_COURSECODE','EXT_MISCINFO','PRIORITYDESCRIPTION','STATUS',
             'SOURCE','SERVICETEAMNAME','COL13','L1_CATEGORY','L2_CATEGORY',
             'ACCOUNTNAME','SERVICE_REQUEST','REGION','COUNTRY','CONTENT_TEAM',
             'PROCESSOR_ID','PROCESSOR_NAME','DOMAIN_CATEGORY','INITIAL_SERVICE_TEAM','IRT_HOURS')

$raw  = Import-Csv $CsvPath -Header $headers | Select-Object -Skip 2 | Where-Object { $_.DISPLAYID -match '^\d' }
$today = [datetime]::Today

# ── Step 3: Resolve processor name ───────────────────────────
function Resolve-Processor($row) {
    if ($row.PROCESSOR_ID -eq 'SAP_GERMANY_USER') {
        $misc = $row.EXT_MISCINFO.Trim() -replace '\s*[-_]?\s*(ILT|Regrant.*|sub-exp.*).*', ''
        $misc = $misc.Trim()
        if ($misc -eq '' -or $misc -eq '(No Value)') { return 'Unassigned (DE)' }
        # Map known short names to full names
        switch ($misc.ToLower()) {
            'mjay'   { return 'Mritunjay Kumar Singh' }
            'shub'   { return 'Shubham Thakur' }
            'ani'    { return 'Anjali Tripathi' }
            'swarna' { return 'Swarna Gupta' }
            'kriti'  { return 'Kriti Bhalla' }
            'naren'  { return 'Narendran Raja' }
            default  { return (Get-Culture).TextInfo.ToTitleCase($misc.ToLower()) }
        }
    }
    $n = $row.PROCESSOR_NAME.Trim()
    if ($n -eq '') { return 'Unassigned' }
    return $n
}

# ── Step 4: Build per-processor stats ────────────────────────
Write-Host "Computing stats..." -ForegroundColor Cyan

$processed = $raw | ForEach-Object {
    $proc    = Resolve-Processor $_
    $created = [datetime]::Parse($_.CREATED_ON_DATE.Trim())
    $ageDays = ($today - $created).Days
    $isActive = $_.STATUS -notin @('Closed','Completed')
    $isMtd    = $_.MONTH -eq 'June'
    [PSCustomObject]@{
        Processor = $proc
        Created   = $created
        AgeDays   = $ageDays
        IsActive  = $isActive
        IsMtd     = $isMtd
        Status    = $_.STATUS
    }
}

$processors = $processed | Group-Object Processor | ForEach-Object {
    $grp         = $_.Group
    $ytdInflow   = $grp.Count
    $active      = ($grp | Where-Object { $_.IsActive }).Count
    $closureRate = if ($ytdInflow -gt 0) { [math]::Round((($ytdInflow - $active) / $ytdInflow) * 100) } else { 0 }
    $mtdInflow   = ($grp | Where-Object { $_.IsMtd }).Count
    $mtdActive   = ($grp | Where-Object { $_.IsMtd -and $_.IsActive }).Count
    $mtdCloses   = $mtdInflow - $mtdActive
    $aged30      = ($grp | Where-Object { $_.IsActive -and $_.AgeDays -gt 30 }).Count
    $aged16to30  = ($grp | Where-Object { $_.IsActive -and $_.AgeDays -ge 16 -and $_.AgeDays -le 30 }).Count

    [PSCustomObject]@{
        Name         = $_.Name
        YtdInflow    = $ytdInflow
        Active       = $active
        ClosureRate  = $closureRate
        MtdInflow    = $mtdInflow
        MtdCloses    = $mtdCloses
        Aged30       = $aged30
        Aged16to30   = $aged16to30
    }
} | Where-Object { $_.YtdInflow -gt 5 } | Sort-Object YtdInflow -Descending

# ── Step 5: Determine award winners ──────────────────────────

# Star of the Week: highest closure rate (min 200 YTD), tiebreak = lowest active
$star = $processors | Where-Object { $_.YtdInflow -ge 200 } |
        Sort-Object ClosureRate -Descending, Active | Select-Object -First 1

# Backlog Buster: zero aged30 cases AND lowest active (min 50 YTD)
$backlogBuster = $processors | Where-Object { $_.YtdInflow -ge 50 -and $_.Aged30 -eq 0 } |
                 Sort-Object Active | Select-Object -First 1

# Most Improved: biggest active reduction proxy = highest MTD closes relative to active (min 50 YTD)
# Since we don't have prior week, use highest MTD closes as proxy
$mostImproved = $processors | Where-Object { $_.YtdInflow -ge 50 } |
                Sort-Object MtdCloses -Descending | Select-Object -First 1

# Top Volume: highest MTD inflow
$topVolume = $processors | Sort-Object MtdInflow -Descending | Select-Object -First 1

# ── Step 6: Build award badge data ───────────────────────────
$starReason = "$($star.ClosureRate)% YTD closure rate with $($star.Active) active cases. Consistently leading on quality and efficiency."

# ── Step 7: Build HTML table rows ────────────────────────────
$tableRows = ""
foreach ($p in $processors) {
    $isStar     = $p.Name -eq $star.Name
    $isBuster   = $backlogBuster -and $p.Name -eq $backlogBuster.Name
    $isImproved = $mostImproved -and $p.Name -eq $mostImproved.Name
    $isVolume   = $p.Name -eq $topVolume.Name

    $nameDisplay = $p.Name
    if ($isStar)     { $nameDisplay += " 🥇" }
    if ($isBuster)   { $nameDisplay += " 🧹" }
    if ($isImproved) { $nameDisplay += " 🚀" }
    if ($isVolume -and -not $isStar -and -not $isBuster -and -not $isImproved) { $nameDisplay += " 🤝" }

    $rowClass = if ($isStar) { ' class="highlight-row"' } else { '' }

    $rateColor = if ($p.ClosureRate -ge 95) { '#007050' } elseif ($p.ClosureRate -ge 90) { '#CC8800' } else { '#C00000' }
    $aged30Display = if ($p.Aged30 -eq 0) { '<span style="color:#888">—</span>' } else { "<span style='font-weight:700;color:#C00000'>$($p.Aged30)</span>" }

    $tableRows += @"
        <tr$rowClass>
          <td>$nameDisplay</td>
          <td>$($p.YtdInflow.ToString('N0'))</td>
          <td>$($p.MtdCloses.ToString('N0'))</td>
          <td>$($p.Active.ToString('N0'))</td>
          <td><span style="color:#888">—</span></td>
          <td style="font-weight:700;color:$rateColor">$($p.ClosureRate)%</td>
          <td>$aged30Display</td>
        </tr>
"@
}

# ── Step 8: Team KPIs ─────────────────────────────────────────
$ytdTotal    = ($processed).Count
$mtdTotal    = ($processed | Where-Object { $_.IsMtd }).Count
$activeTotal = ($processed | Where-Object { $_.IsActive }).Count
$closureAll  = [math]::Round((($ytdTotal - $activeTotal) / $ytdTotal) * 100)

# ── Step 9: Write HTML ────────────────────────────────────────
$outputPath = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\GTS_Winners_Circle_$($WeekLabel)_$($WeekRange -replace '[^a-zA-Z0-9]','-').html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>GTS Winners Circle — $WeekRange</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: linear-gradient(135deg, #0a1628 0%, #1a2f5e 100%); min-height: 100vh; padding: 30px 20px; }
  .bulletin { max-width: 900px; margin: 0 auto; }

  .header { background: linear-gradient(135deg, #007DB8, #005a8a); border-radius: 16px; padding: 30px 36px; text-align: center; margin-bottom: 20px; position: relative; overflow: hidden; }
  .header::before { content:''; position:absolute; top:-40px; right:-40px; width:160px; height:160px; background:rgba(255,255,255,0.05); border-radius:50%; }
  .header-trophy { font-size: 48px; margin-bottom: 8px; }
  .header h1 { font-size: 28px; font-weight: 800; color: #fff; letter-spacing: 1px; }
  .header h2 { font-size: 15px; font-weight: 400; color: rgba(255,255,255,0.8); margin-top: 6px; }
  .header-tagline { font-size: 13px; color: rgba(255,255,255,0.6); margin-top: 8px; font-style: italic; }

  .snapshot { background: #fff; border-radius: 14px; padding: 20px 24px; margin-bottom: 20px; }
  .section-title { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #007DB8; margin-bottom: 14px; }
  .kpi-row { display: flex; gap: 12px; }
  .kpi-box { flex: 1; background: #EAF4FB; border-radius: 10px; padding: 14px; text-align: center; border-top: 3px solid #007DB8; }
  .kpi-label { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi-val   { font-size: 24px; font-weight: 800; color: #007DB8; margin-top: 4px; }
  .kpi-sub   { font-size: 10px; color: #999; margin-top: 2px; }

  .star-card { background: linear-gradient(135deg, #FFD700, #FFA500); border-radius: 14px; padding: 24px 30px; margin-bottom: 20px; display: flex; align-items: center; gap: 20px; position: relative; overflow: hidden; }
  .star-card::after { content: '⭐'; position: absolute; right: 20px; top: 10px; font-size: 60px; opacity: 0.15; }
  .star-emoji { font-size: 52px; }
  .star-label  { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: rgba(0,0,0,0.5); }
  .star-name   { font-size: 26px; font-weight: 800; color: #1a1a1a; margin-top: 2px; }
  .star-reason { font-size: 13px; color: #333; margin-top: 6px; }
  .star-stats  { display: flex; gap: 16px; margin-top: 10px; }
  .star-stat   { background: rgba(0,0,0,0.1); border-radius: 6px; padding: 5px 12px; font-size: 11px; font-weight: 600; color: #222; }

  .winners-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }
  .winner-card { background: #fff; border-radius: 12px; padding: 18px 16px; text-align: center; border-top: 4px solid #007DB8; }
  .winner-card.backlog  { border-color: #C00000; }
  .winner-card.improved { border-color: #00A86B; }
  .winner-card.csat     { border-color: #9B59B6; }
  .winner-card.special  { border-color: #E67E22; }
  .winner-emoji { font-size: 30px; margin-bottom: 8px; }
  .winner-badge { font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; padding: 3px 8px; border-radius: 20px; margin-bottom: 8px; display: inline-block; }
  .backlog  .winner-badge { background: #FDECEA; color: #C00000; }
  .improved .winner-badge { background: #E2EFDA; color: #007050; }
  .csat     .winner-badge { background: #F0E6FF; color: #7D3C98; }
  .special  .winner-badge { background: #FEF0E6; color: #C0392B; }
  .winner-name   { font-size: 15px; font-weight: 800; color: #222; margin: 6px 0 4px; }
  .winner-metric { font-size: 22px; font-weight: 800; margin: 4px 0; }
  .backlog  .winner-metric { color: #C00000; }
  .improved .winner-metric { color: #00A86B; }
  .csat     .winner-metric { color: #9B59B6; }
  .special  .winner-metric { color: #E67E22; }
  .winner-desc  { font-size: 11px; color: #666; line-height: 1.5; margin-top: 4px; }
  .winner-quote { font-size: 11px; color: #555; font-style: italic; margin-top: 6px; line-height: 1.4; }
  .badge-new { display: inline-block; background: #FFD700; color: #000; font-size: 9px; font-weight: 800; padding: 2px 6px; border-radius: 4px; margin-left: 4px; }

  .table-card { background: #fff; border-radius: 14px; padding: 20px 24px; margin-bottom: 20px; }
  .ind-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 12px; }
  .ind-table th { background: #007DB8; color: #fff; font-weight: 700; padding: 9px 12px; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; }
  .ind-table td { padding: 9px 12px; border-bottom: 1px solid #f0f0f0; }
  .ind-table tr:last-child td { border-bottom: none; }
  .ind-table tr:nth-child(even) td { background: #fafafa; }
  .highlight-row td { background: #FFF9E6 !important; font-weight: 600; }

  .footer-card { background: linear-gradient(135deg, #1a2f5e, #0a1628); border-radius: 14px; padding: 22px 28px; color: #fff; display: flex; gap: 30px; align-items: flex-start; }
  .footer-section { flex: 1; }
  .footer-section h4 { font-size: 11px; text-transform: uppercase; letter-spacing: 1.5px; color: rgba(255,255,255,0.5); margin-bottom: 10px; }
  .goal-item { font-size: 13px; color: #fff; margin-bottom: 6px; }
  .tip-text  { font-size: 13px; color: rgba(255,255,255,0.8); font-style: italic; line-height: 1.5; }
  .footer-divider { width: 1px; background: rgba(255,255,255,0.1); }
  .confetti-wrap { position: absolute; top: 0; left: 0; width: 100%; height: 100%; pointer-events: none; overflow: hidden; }
  .c { position: absolute; top: -12px; opacity: 0; animation: fall linear infinite; border-radius: 2px; }
  @keyframes fall {
    0%   { transform: translateY(0) rotate(0deg); opacity: 1; }
    100% { transform: translateY(280px) rotate(720deg); opacity: 0; }
  }
</style>
</head>
<body>
<div class="bulletin">

  <div class="header">
    <div class="confetti-wrap" id="confetti"></div>
    <div class="header-trophy">🏆</div>
    <h1>GTS Winners — Week of $WeekRange</h1>
    <h2>GTS Team Bulletin</h2>
    <div class="header-tagline">Your team. Your wins. Keep pushing! 🔥</div>
  </div>

  <div class="snapshot">
    <div class="section-title">📊 Team Snapshot</div>
    <div class="kpi-row">
      <div class="kpi-box"><div class="kpi-label">YTD Inflow</div><div class="kpi-val">$($ytdTotal.ToString('N0'))</div><div class="kpi-sub">Jan–Jun</div></div>
      <div class="kpi-box"><div class="kpi-label">MTD Inflow</div><div class="kpi-val">$($mtdTotal.ToString('N0'))</div><div class="kpi-sub">June so far</div></div>
      <div class="kpi-box"><div class="kpi-label">Active Cases</div><div class="kpi-val">$($activeTotal.ToString('N0'))</div><div class="kpi-sub">As of today</div></div>
      <div class="kpi-box"><div class="kpi-label">Closure Rate</div><div class="kpi-val">$($closureAll)%</div><div class="kpi-sub">YTD</div></div>
    </div>
  </div>

  <div class="star-card">
    <div class="star-emoji">🥇</div>
    <div>
      <div class="star-label">⭐ Star of the Week</div>
      <div class="star-name">$($star.Name)</div>
      <div class="star-reason">$starReason</div>
      <div class="star-stats">
        <div class="star-stat">📥 YTD Inflow: $($star.YtdInflow.ToString('N0'))</div>
        <div class="star-stat">✅ YTD Closed: $(($star.YtdInflow - $star.Active).ToString('N0'))</div>
        <div class="star-stat">📊 Rate: $($star.ClosureRate)%</div>
        <div class="star-stat">📂 Active: $($star.Active)</div>
      </div>
    </div>
  </div>

  <div class="winners-grid">
    <div class="winner-card backlog">
      <div class="winner-emoji">🧹</div>
      <div class="winner-badge">Backlog Buster</div>
      <div class="winner-name">$($backlogBuster.Name)</div>
      <div class="winner-metric">0 aged cases</div>
      <div class="winner-desc">Zero cases older than 30 days with $($backlogBuster.Active) active cases. Keeping the backlog clean while staying on top of inflow. 🧹</div>
    </div>
    <div class="winner-card improved">
      <div class="winner-emoji">🚀</div>
      <div class="winner-badge">Volume Leader</div>
      <div class="winner-name">$($mostImproved.Name)</div>
      <div class="winner-metric">$($mostImproved.MtdCloses) MTD closes</div>
      <div class="winner-desc">Highest case closures in June so far — leading the team in throughput. Outstanding push! 💪</div>
    </div>
    <div class="winner-card csat">
      <div class="winner-emoji">⭐</div>
      <div class="winner-badge">CSAT Champion</div>
      <div class="winner-name">—</div>
      <div class="winner-metric">Pending</div>
      <div class="winner-desc">CSAT data not yet available for this period.</div>
    </div>
    <div class="winner-card special">
      <div class="winner-emoji">🤝</div>
      <div class="winner-badge">Top Inflow <span class="badge-new">THIS WEEK</span></div>
      <div class="winner-name">$($topVolume.Name)</div>
      <div class="winner-metric">$($topVolume.MtdInflow) MTD cases</div>
      <div class="winner-desc">Highest case volume received in June — keeping pace and handling the load. 🎯</div>
    </div>
  </div>

  <div class="table-card">
    <div class="section-title">📈 Individual Snapshot — As of $RunDate</div>
    <table class="ind-table">
      <thead>
        <tr>
          <th>Agent</th><th>YTD Inflow</th><th>MTD Closes</th>
          <th>Active</th><th>vs Yesterday</th><th>Closure Rate</th><th>&gt;30d</th>
        </tr>
      </thead>
      <tbody>
        $tableRows
      </tbody>
    </table>
  </div>

  <div style="background:#D6F0E0;border-radius:14px;padding:20px 24px;margin-bottom:20px;">
    <div style="margin-bottom:14px;">
      <span style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#007050;">🏆 What Success Looks Like</span>
    </div>
    <div style="display:flex;gap:12px;margin-top:4px;">
      <div style="flex:1;background:#F0E6FF;border-radius:10px;padding:14px 10px;text-align:center;border-top:3px solid #7D3C98;">
        <div style="font-size:22px;margin-bottom:6px;">📅</div>
        <div style="font-size:11px;font-weight:800;color:#7D3C98;line-height:1.3;">Cases are FRESH</div>
        <div style="font-size:10px;color:#555;margin-top:4px;">90% within 0–15 days</div>
      </div>
      <div style="flex:1;background:#F0E6FF;border-radius:10px;padding:14px 10px;text-align:center;border-top:3px solid #7D3C98;">
        <div style="font-size:22px;margin-bottom:6px;">🚀</div>
        <div style="font-size:11px;font-weight:800;color:#7D3C98;line-height:1.3;">We close what we receive</div>
        <div style="font-size:10px;color:#555;margin-top:4px;">Closure Rate ≥ 90%</div>
      </div>
      <div style="flex:1;background:#F0E6FF;border-radius:10px;padding:14px 10px;text-align:center;border-top:3px solid #7D3C98;">
        <div style="font-size:22px;margin-bottom:6px;">👻</div>
        <div style="font-size:11px;font-weight:800;color:#7D3C98;line-height:1.3;">No ghost cases. Ever.</div>
        <div style="font-size:10px;color:#555;margin-top:4px;">Zero idle cases</div>
      </div>
      <div style="flex:1;background:#F0E6FF;border-radius:10px;padding:14px 10px;text-align:center;border-top:3px solid #7D3C98;">
        <div style="font-size:22px;margin-bottom:6px;">🔥</div>
        <div style="font-size:11px;font-weight:800;color:#7D3C98;line-height:1.3;">No case left behind.</div>
        <div style="font-size:10px;color:#555;margin-top:4px;">Zero cases &gt;30 days</div>
      </div>
      <div style="flex:1;background:#F0E6FF;border-radius:10px;padding:14px 10px;text-align:center;border-top:3px solid #7D3C98;">
        <div style="font-size:22px;margin-bottom:6px;">⭐</div>
        <div style="font-size:11px;font-weight:800;color:#7D3C98;line-height:1.3;">Customers say WOW.</div>
        <div style="font-size:10px;color:#555;margin-top:4px;">CSAT reflects fast resolution</div>
      </div>
    </div>
  </div>

  <div class="footer-card">
    <div class="footer-section">
      <h4>🎯 This Week's Goal</h4>
      <div class="goal-item">✅ Continue clearing &gt;30-day cases — keep momentum on backlog reduction</div>
      <div class="goal-item">✅ Maintain 90%+ closure rate</div>
      <div class="goal-item">✅ Keep new cases moving — no idle cases</div>
    </div>
    <div class="footer-divider"></div>
    <div class="footer-section">
      <h4>💡 Tip of the Week</h4>
      <div class="tip-text">$TipOfWeek</div>
    </div>
    <div class="footer-divider"></div>
    <div class="footer-section">
      <h4>📅 Next Bulletin</h4>
      <div class="tip-text">Next GTS Winners Circle drops <strong>$NextBulletin</strong>. Who will be Star of the Week? Keep pushing! 💪</div>
    </div>
  </div>

</div>
<script>
  const colors = ['#FFD700','#FF6B6B','#4ECDC4','#45B7D1','#96CEB4','#FFEAA7','#DDA0DD','#98D8C8','#FF8C42','#A8E6CF'];
  const wrap = document.getElementById('confetti');
  for (let i = 0; i < 40; i++) {
    const c = document.createElement('div');
    c.className = 'c';
    c.style.left = Math.random() * 100 + '%';
    c.style.background = colors[Math.floor(Math.random() * colors.length)];
    c.style.width = (6 + Math.random() * 8) + 'px';
    c.style.height = (6 + Math.random() * 8) + 'px';
    c.style.animationDuration = (2 + Math.random() * 3) + 's';
    c.style.animationDelay = (Math.random() * 3) + 's';
    c.style.borderRadius = Math.random() > 0.5 ? '50%' : '2px';
    wrap.appendChild(c);
  }
</script>
</body>
</html>
"@

$html | Out-File $outputPath -Encoding UTF8
Write-Host ""
Write-Host "Done! Output saved to:" -ForegroundColor Green
Write-Host "  $outputPath" -ForegroundColor Yellow
Start-Process $outputPath
