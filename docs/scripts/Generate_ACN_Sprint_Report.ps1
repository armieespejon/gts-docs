# ============================================================
# GTS ACN Cert Week Sprint Report - Generator
# Update the CONFIG block below, then run:
#   powershell -ExecutionPolicy Bypass -File "Generate_ACN_Sprint_Report.ps1"
# ============================================================

# ?? CONFIG ???????????????????????????????????????????????????
$OutputDir   = "C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie"
$ReportDate  = "June 17, 2026"
$SprintLabel = "June 1-7, 2026"

# Program-level KPIs (header chips)
$TotalAttempts   = "12,000"
$LearnerseCertified = "~8,200"
$PassRate        = "68%"
$TicketsToSAP    = "1,636"

# Certifications Issued
$CertAIG   = "4,904"
$CertBCBAI = "2,911"
$CertBCSBS = "378"
$CertTotal = "8,193"

# SBA C_AIG KPI tiles
$SBA_Total      = "1,162"
$SBA_Resolved   = "1,136"
$SBA_ResPct     = "97.8%"
$SBA_Open       = "26"
$SBA_Days       = "7"

# Issue drivers (percentages)
$Driver_AICore   = "70.5"
$Driver_Score    = "6.0"
$Driver_Other    = "23.5"

# Daily volume (Jun 1-7)
$Day1 = 45;  $Day2 = 27;  $Day3 = 53
$Day4 = 229; $Day5 = 152; $Day6 = 437; $Day7 = 219
$PeakDay   = "Jun 6"
$PeakCount = "437"

# CSAT
$CSAT_Surveys      = "70"
$CSAT_Satisfied    = "55"
$CSAT_SatisfiedPct = "79%"
$CSAT_Neutral      = "9"
$CSAT_NeutralPct   = "13%"
$CSAT_Dissatisfied = "6"
$CSAT_DissPct      = "9%"
$CSAT_CoverageOf   = "70 of 1,162"
$CSAT_CoveragePct  = "6%"

# Sprint extension notice (set to $true to show, $false to hide)
$ShowSprintExtension = $true
$SprintExtensionText = "The sprint has been extended to end of June to allow Accenture to close the gap between their Cert Week results and the forecast provided the week prior. Target numbers are still under discussion."

# AI Core status note
$AICoreFix = "Hotfixes applied. Development team working on permanent fix - ETA end of July 2026."

# ?? END CONFIG ???????????????????????????????????????????????

# Compute SVG chart points (max scale 500, y = 140 - (val/500)*130)
function Get-Y($val) { return [math]::Round(140 - ($val / 500 * 130)) }
$pts = "40,$(Get-Y $Day1) 110,$(Get-Y $Day2) 180,$(Get-Y $Day3) 250,$(Get-Y $Day4) 320,$(Get-Y $Day5) 390,$(Get-Y $Day6) 460,$(Get-Y $Day7)"
$y1 = Get-Y $Day1; $y2 = Get-Y $Day2; $y3 = Get-Y $Day3
$y4 = Get-Y $Day4; $y5 = Get-Y $Day5; $y6 = Get-Y $Day6; $y7 = Get-Y $Day7
$lbl1y = $y1 - 7; $lbl2y = $y2 - 7; $lbl3y = $y3 - 7
$lbl4y = $y4 - 7; $lbl5y = $y5 - 7; $lbl6y = $y6 - 7; $lbl7y = $y7 - 7

# Sprint extension block
$sprintExtHtml = ""
if ($ShowSprintExtension) {
    $sprintExtHtml = @"
<!-- SPRINT EXTENSION NOTICE -->
<div style="margin-bottom:28px;padding:12px 18px;background:#fffbeb;border-radius:8px;border:1px solid #fcd34d;border-left:4px solid #b45309;display:flex;align-items:flex-start;gap:12px;">
  <div style="font-size:18px;line-height:1;">&#128204;</div>
  <div>
    <div style="font-size:12px;font-weight:700;color:#92400e;margin-bottom:3px;">Sprint Extended - Until End of June</div>
    <div style="font-size:11px;color:#78350f;line-height:1.6;">$SprintExtensionText</div>
  </div>
</div>
"@
}

$outputFile = "$OutputDir\GTS_ACN_CertWeek_Sprint_Jun1-7.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Accenture x SAP Certification Week -- GTS Support Report</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f8f7f4; color: #1e293b; padding: 36px 32px; max-width: 1140px; margin: 0 auto; }
    .page-header { margin-bottom: 32px; padding: 24px 28px 20px; background: #1e3a5f; border-radius: 12px; }
    .header-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px; }
    .header-top h1 { font-size: 22px; font-weight: 700; color: #fff; }
    .header-top h1 span { color: #fbbf24; }
    .header-meta { font-size: 12px; color: #93c5fd; text-align: right; line-height: 1.6; }
    .header-chips { display: flex; gap: 10px; flex-wrap: wrap; }
    .chip { background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.2); border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #e0f0ff; display: flex; align-items: center; gap: 6px; }
    .chip .dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
    .section { margin-bottom: 36px; }
    .section-header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; }
    .section-header h2 { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.4px; color: #fff; white-space: nowrap; background: #1e3a5f; padding: 6px 16px; border-radius: 4px; }
    .section-rule { flex: 1; height: 2px; background: linear-gradient(to right, #1e3a5f, transparent); }
    .cert-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }
    .cert-card { background: var(--bg); border-radius: 8px; padding: 16px 16px 14px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.12); }
    .cert-card.c1 { --bg: #1d4ed8; } .cert-card.c2 { --bg: #b45309; } .cert-card.c3 { --bg: #047857; } .cert-card.total { --bg: #1e3a5f; }
    .cert-code { font-size: 13px; font-weight: 700; color: rgba(255,255,255,0.85); margin-bottom: 6px; }
    .cert-name { font-size: 10px; color: rgba(255,255,255,0.6); margin-bottom: 10px; line-height: 1.4; }
    .cert-num { font-size: 28px; font-weight: 700; color: #fff; }
    .sba-kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 24px; }
    .sba-tile { background: #fff; border-radius: 8px; padding: 18px 16px; border-bottom: 4px solid var(--tc); box-shadow: 0 1px 3px rgba(0,0,0,0.07); }
    .sba-tile.t1 { --tc: #1d4ed8; } .sba-tile.t2 { --tc: #047857; } .sba-tile.t3 { --tc: #b91c1c; } .sba-tile.t4 { --tc: #64748b; }
    .sba-tile-val { font-size: 32px; font-weight: 700; color: var(--tc); line-height: 1; margin-bottom: 4px; }
    .sba-tile-label { font-size: 12px; font-weight: 700; color: #1e293b; margin-bottom: 3px; }
    .sba-tile-sub { font-size: 11px; color: #94a3b8; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
    .panel { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    .panel-title { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #64748b; margin-bottom: 16px; }
    .driver-item { margin-bottom: 14px; }
    .driver-item:last-child { margin-bottom: 0; }
    .driver-header { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 5px; }
    .driver-name { font-size: 13px; font-weight: 600; color: #1e293b; }
    .driver-pct { font-size: 13px; font-weight: 700; color: #1d4ed8; }
    .driver-track { background: #e2e8f0; border-radius: 4px; height: 10px; overflow: hidden; }
    .driver-fill { height: 100%; border-radius: 4px; }
    .driver-note { font-size: 10px; color: #94a3b8; margin-top: 3px; }
    .chart-svg-wrap { background: #fff; border-radius: 10px; padding: 20px 20px 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    .chart-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #64748b; margin-bottom: 12px; }
    .obs-item { padding: 12px 14px; border-radius: 8px; background: #f0f6ff; border-left: 4px solid #1d4ed8; margin-bottom: 10px; }
    .obs-item:last-child { margin-bottom: 0; }
    .obs-title { font-size: 12px; font-weight: 700; color: #1e293b; margin-bottom: 4px; }
    .obs-body { font-size: 11px; color: #475569; line-height: 1.6; }
    .csat-tiles { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 24px; }
    .csat-tile { background: #fff; border-radius: 8px; padding: 14px 12px; text-align: center; border-top: 4px solid var(--ct); box-shadow: 0 1px 3px rgba(0,0,0,0.07); }
    .csat-tile.ct1 { --ct: #1d4ed8; } .csat-tile.ct2 { --ct: #047857; } .csat-tile.ct3 { --ct: #b45309; } .csat-tile.ct4 { --ct: #b91c1c; } .csat-tile.ct5 { --ct: #1e3a5f; }
    .csat-val { font-size: 26px; font-weight: 700; color: var(--ct); }
    .csat-val span { font-size: 14px; }
    .csat-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #64748b; margin-top: 4px; }
    .csat-sub { font-size: 10px; color: #94a3b8; margin-top: 3px; }
    .voc-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; }
    .voc-col-header { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; padding: 7px 12px; border-radius: 4px; margin-bottom: 10px; text-align: center; }
    .voc-col-header.pos { background: #dcfce7; color: #14532d; } .voc-col-header.neu { background: #fef9c3; color: #713f12; } .voc-col-header.neg { background: #fee2e2; color: #7f1d1d; }
    .voc-quote { background: #f8f7f4; border-radius: 6px; padding: 10px 12px; margin-bottom: 8px; font-size: 11px; color: #334155; line-height: 1.6; border-left: 3px solid var(--qc); }
    .voc-quote.pos { --qc: #16a34a; } .voc-quote.neu { --qc: #ca8a04; } .voc-quote.neg { --qc: #dc2626; }
    .voc-quote::before { content: '\201C'; font-size: 18px; color: var(--qc); line-height: 0; vertical-align: -4px; margin-right: 3px; }
    .learning-item { display: flex; gap: 16px; padding: 16px 18px; background: #fff; border-radius: 10px; margin-bottom: 10px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .learning-item:last-child { margin-bottom: 0; }
    .learning-num { width: 30px; height: 30px; border-radius: 50%; background: #1e3a5f; color: #fff; font-size: 13px; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 2px; }
    .learning-body { flex: 1; }
    .learning-title { font-size: 13px; font-weight: 700; color: #1e293b; margin-bottom: 4px; }
    .learning-issue { font-size: 11px; color: #64748b; line-height: 1.6; margin-bottom: 6px; }
    .learning-action { font-size: 11px; color: #1d4ed8; font-weight: 600; line-height: 1.5; }
    .learning-action::before { content: 'Action: '; color: #b45309; }
    .footer { margin-top: 16px; font-size: 11px; color: #94a3b8; text-align: center; padding-top: 14px; border-top: 1px solid #e2e8f0; }
  </style>
</head>
<body>

<div class="page-header">
  <div class="header-top">
    <div>
      <h1>Accenture x SAP <span>Certification Week</span> -- GTS Support Report</h1>
      <div style="font-size:13px;color:#93c5fd;margin-top:4px;">Sprint: $SprintLabel &nbsp;.&nbsp; C_AIG . C_BCBAI . C_BCSBS &nbsp;.&nbsp; Partner: Accenture</div>
    </div>
    <div class="header-meta">Prepared by GTS Management<br>As of $ReportDate</div>
  </div>
  <div class="header-chips">
    <div class="chip"><div class="dot" style="background:#7dd3fc"></div> $TotalAttempts Certification Attempts</div>
    <div class="chip"><div class="dot" style="background:#6ee7b7"></div> $LearnerseCertified Learners Certified (~$PassRate pass rate)</div>
    <div class="chip"><div class="dot" style="background:#fde68a"></div> $TicketsToSAP Tickets to SAP</div>
    <div class="chip"><div class="dot" style="background:#c4b5fd"></div> 7-day sprint . Mon-Sun</div>
  </div>
</div>

$sprintExtHtml

<div class="section">
  <div class="section-header"><h2>Certifications Issued -- June 2026</h2><div class="section-rule"></div></div>
  <div class="cert-grid">
    <div class="cert-card c1"><div class="cert-code">C_AIG</div><div class="cert-name">SAP Certified Generative AI Developer</div><div class="cert-num">$CertAIG</div></div>
    <div class="cert-card c2"><div class="cert-code">C_BCBAI</div><div class="cert-name">SAP Certified - Positioning Business AI Solutions</div><div class="cert-num">$CertBCBAI</div></div>
    <div class="cert-card c3"><div class="cert-code">C_BCSBS</div><div class="cert-name">SAP Certified - Positioning the Autonomous Enterprise</div><div class="cert-num">$CertBCSBS</div></div>
    <div class="cert-card total"><div class="cert-code" style="color:#7dd3fc;">TOTAL</div><div class="cert-name">All 3 Certification Tracks</div><div class="cert-num" style="font-size:32px;">$CertTotal</div></div>
  </div>
</div>

<div class="section">
  <div class="section-header"><h2>System Based Exam (SBA) -- C_AIG Support Overview</h2><div class="section-rule"></div></div>
  <div style="font-size:12px;color:#64748b;margin-bottom:16px;">C_AIG -- SAP Certified Generative AI Developer &nbsp;.&nbsp; Sprint: $SprintLabel</div>
  <div class="sba-kpi-row">
    <div class="sba-tile t1"><div class="sba-tile-val">$SBA_Total</div><div class="sba-tile-label">Total Cases</div><div class="sba-tile-sub">C_AIG SBA</div></div>
    <div class="sba-tile t2"><div class="sba-tile-val">$SBA_Resolved</div><div class="sba-tile-label">Resolved</div><div class="sba-tile-sub">$SBA_ResPct resolution rate</div></div>
    <div class="sba-tile t3"><div class="sba-tile-val">$SBA_Open</div><div class="sba-tile-label">Still Open</div><div class="sba-tile-sub">Score / reschedule / pending</div></div>
    <div class="sba-tile t4"><div class="sba-tile-val">$SBA_Days</div><div class="sba-tile-label">Sprint Days</div><div class="sba-tile-sub">Mon-Sun</div></div>
  </div>

  <div class="grid-2">
    <div class="panel">
      <div class="panel-title">Top Issue Drivers</div>
      <div class="driver-item">
        <div class="driver-header"><span class="driver-name">AI Core Issues</span><span class="driver-pct">$Driver_AICore%</span></div>
        <div class="driver-track"><div class="driver-fill" style="width:$($Driver_AICore)%;background:#1d4ed8;"></div></div>
        <div class="driver-note">AI API Error + Network/Technical Error + Template/Prompt Editor Issue</div>
        <div style="margin-top:6px;padding:6px 10px;background:#eff6ff;border-radius:4px;font-size:10px;color:#1d4ed8;border-left:3px solid #1d4ed8;"><strong>Status:</strong> $AICoreFix</div>
      </div>
      <div class="driver-item">
        <div class="driver-header"><span class="driver-name">Score / Result Inquiry</span><span class="driver-pct">$Driver_Score%</span></div>
        <div class="driver-track"><div class="driver-fill" style="width:$($Driver_Score)%;background:#b45309;"></div></div>
        <div class="driver-note">Score review, result disputes, retry &amp; attempt management</div>
      </div>
      <div class="driver-item">
        <div class="driver-header"><span class="driver-name">Other</span><span class="driver-pct">$Driver_Other%</span></div>
        <div class="driver-track"><div class="driver-fill" style="width:$($Driver_Other)%;background:#64748b;"></div></div>
        <div class="driver-note">Unclassified + How-To &amp; User Guidance + Access &amp; Enrollment + Task/Exercise Issues + Schedule Cancellation Requests</div>
      </div>
      <div style="margin-top:16px;padding:10px 12px;background:#f0f6ff;border-radius:6px;font-size:11px;color:#475569;border-left:4px solid #1d4ed8;">
        AI Core Issues account for <strong style="color:#1d4ed8;">$Driver_AICore%</strong> of all cases -- rooted in SAP AI Core product. Frequency escalated from Jun 4 onwards.
      </div>
    </div>

    <div>
      <div class="chart-svg-wrap" style="margin-bottom:14px;">
        <div class="chart-label">Daily Case Volume -- $SprintLabel</div>
        <svg viewBox="0 0 480 160" style="width:100%;height:160px;">
          <line x1="40" y1="10"  x2="460" y2="10"  stroke="#e5e7eb" stroke-width="1"/>
          <line x1="40" y1="45"  x2="460" y2="45"  stroke="#e5e7eb" stroke-width="1"/>
          <line x1="40" y1="80"  x2="460" y2="80"  stroke="#e5e7eb" stroke-width="1"/>
          <line x1="40" y1="115" x2="460" y2="115" stroke="#e5e7eb" stroke-width="1"/>
          <line x1="40" y1="140" x2="460" y2="140" stroke="#e5e7eb" stroke-width="1"/>
          <text x="36" y="14"  text-anchor="end" font-size="9" fill="#9ca3af">500</text>
          <text x="36" y="49"  text-anchor="end" font-size="9" fill="#9ca3af">375</text>
          <text x="36" y="84"  text-anchor="end" font-size="9" fill="#9ca3af">250</text>
          <text x="36" y="119" text-anchor="end" font-size="9" fill="#9ca3af">125</text>
          <text x="36" y="144" text-anchor="end" font-size="9" fill="#9ca3af">0</text>
          <polyline points="$pts" fill="none" stroke="#1d4ed8" stroke-width="2.5" stroke-linejoin="round"/>
          <circle cx="40"  cy="$y1" r="4" fill="#1d4ed8"/>
          <circle cx="110" cy="$y2" r="4" fill="#1d4ed8"/>
          <circle cx="180" cy="$y3" r="4" fill="#1d4ed8"/>
          <circle cx="250" cy="$y4" r="4" fill="#1d4ed8"/>
          <circle cx="320" cy="$y5" r="4" fill="#1d4ed8"/>
          <circle cx="390" cy="$y6" r="5" fill="#b91c1c"/>
          <circle cx="460" cy="$y7" r="4" fill="#1d4ed8"/>
          <text x="40"  y="$lbl1y" text-anchor="middle" font-size="9" fill="#64748b">$Day1</text>
          <text x="110" y="$lbl2y" text-anchor="middle" font-size="9" fill="#64748b">$Day2</text>
          <text x="180" y="$lbl3y" text-anchor="middle" font-size="9" fill="#64748b">$Day3</text>
          <text x="250" y="$lbl4y" text-anchor="middle" font-size="9" fill="#64748b">$Day4</text>
          <text x="320" y="$lbl5y" text-anchor="middle" font-size="9" fill="#64748b">$Day5</text>
          <text x="390" y="$lbl6y" text-anchor="middle" font-size="10" fill="#b91c1c" font-weight="bold">$Day6</text>
          <text x="460" y="$lbl7y" text-anchor="middle" font-size="9" fill="#64748b">$Day7</text>
          <text x="40"  y="155" text-anchor="middle" font-size="9" fill="#9ca3af">Jun 1</text>
          <text x="110" y="155" text-anchor="middle" font-size="9" fill="#9ca3af">Jun 2</text>
          <text x="180" y="155" text-anchor="middle" font-size="9" fill="#9ca3af">Jun 3</text>
          <text x="250" y="155" text-anchor="middle" font-size="9" fill="#9ca3af">Jun 4</text>
          <text x="320" y="155" text-anchor="middle" font-size="9" fill="#9ca3af">Jun 5</text>
          <text x="390" y="155" text-anchor="middle" font-size="9" fill="#b91c1c" font-weight="bold">$PeakDay</text>
          <text x="460" y="155" text-anchor="middle" font-size="9" fill="#9ca3af">Jun 7</text>
        </svg>
      </div>
      <div class="panel">
        <div class="panel-title">Key Observations</div>
        <div class="obs-item">
          <div class="obs-title">Peak Day -- $PeakDay ($PeakCount cases)</div>
          <div class="obs-body">$PeakCount cases on Saturday -- highest single-day volume. Weekend backlog from learners who attempted exams Friday evening drove the spike.</div>
        </div>
        <div class="obs-item" style="border-left-color:#047857;background:#f0fdf4;">
          <div class="obs-title">Strong Resolution -- $SBA_ResPct</div>
          <div class="obs-body">$SBA_Resolved of $SBA_Total cases resolved within the sprint window. $SBA_Open cases remain open -- score inquiries, reschedules, and pending learner action.</div>
        </div>
        <div class="obs-item" style="border-left-color:#b45309;background:#fffbeb;">
          <div class="obs-title">AI Core Product -- Root Cause</div>
          <div class="obs-body">$Driver_AICore% of cases rooted in SAP AI Core. Frequency escalated significantly from Thu Jun 4 onwards, driving the bulk of volume.</div>
        </div>
      </div>
    </div>
  </div>
</div>

<div class="section">
  <div class="section-header"><h2>CSAT -- C_AIG Survey Results</h2><div class="section-rule"></div></div>
  <div class="csat-tiles">
    <div class="csat-tile ct1"><div class="csat-val">$CSAT_Surveys</div><div class="csat-label">Surveys Received</div><div class="csat-sub">&nbsp;</div></div>
    <div class="csat-tile ct2"><div class="csat-val">$CSAT_Satisfied <span>($CSAT_SatisfiedPct)</span></div><div class="csat-label">Satisfied</div><div class="csat-sub">Top-2-Box</div></div>
    <div class="csat-tile ct3"><div class="csat-val">$CSAT_Neutral</div><div class="csat-label">Neutral</div><div class="csat-sub">$CSAT_NeutralPct</div></div>
    <div class="csat-tile ct4"><div class="csat-val">$CSAT_Dissatisfied <span>($CSAT_DissPct)</span></div><div class="csat-label">Dissatisfied</div><div class="csat-sub">&nbsp;</div></div>
    <div class="csat-tile ct5"><div class="csat-val" style="font-size:18px;padding-top:4px;">$CSAT_CoverageOf</div><div class="csat-label">CSAT Coverage</div><div class="csat-sub">$CSAT_CoveragePct of cases received a survey</div></div>
  </div>
  <div class="voc-grid">
    <div>
      <div class="voc-col-header pos">Positive</div>
      <div class="voc-quote pos">Quickly resolved by team. So satisfied with the instant support.</div>
      <div class="voc-quote pos">There was a connectivity/system issue when I was in the certification exam. Due to the technical error, SAP kindly allowed me to sit the exam again. Thanks for your support.</div>
      <div class="voc-quote pos">Though creation of the support case was delayed due to issues with the portal, resolution was provided quite fast once the case was created.</div>
      <div class="voc-quote pos">I got help fast and it was very helpful. Thanks!</div>
    </div>
    <div>
      <div class="voc-col-header neu">Neutral</div>
      <div class="voc-quote neu">Luckily I was able to complete the certification despite the recurring error. The resolution was provided quite late.</div>
      <div class="voc-quote neu">Solution can be provided a bit faster.</div>
    </div>
    <div>
      <div class="voc-col-header neg">Dissatisfied</div>
      <div class="voc-quote neg">It is too frustrating -- you need to focus more on technical issues than exam.</div>
      <div class="voc-quote neg">Always facing so many issues while doing certification and losing hope as well.</div>
      <div class="voc-quote neg">I was expecting some kind of fix -- I made a mistake saving and was given a 0. Will have to try again but hoped SAP team would have been more understanding.</div>
    </div>
  </div>
</div>

<div class="section">
  <div class="section-header"><h2>Learnings -- What to Do Differently Next Time</h2><div class="section-rule"></div></div>
  <div class="learning-item">
    <div class="learning-num">1</div>
    <div class="learning-body">
      <div class="learning-title">SLA Expectation Setting</div>
      <div class="learning-issue">Learners expected real-time replies and kept updating tickets when they didn't hear back, creating unnecessary follow-up cases.</div>
      <div class="learning-action">Partner shares SLA expectations and FAQ before go-live. Set clear turnaround time expectations upfront.</div>
    </div>
  </div>
  <div class="learning-item">
    <div class="learning-num">2</div>
    <div class="learning-body">
      <div class="learning-title">Exam Scheduling Distribution</div>
      <div class="learning-issue">Weekend exam clustering caused 900+ Monday ticket spikes that overwhelmed the support queue.</div>
      <div class="learning-action">Incentivize weekday exams. Share scheduling guidance earlier and reinforce with partner program managers.</div>
    </div>
  </div>
  <div class="learning-item">
    <div class="learning-num">3</div>
    <div class="learning-body">
      <div class="learning-title">Ticket Consolidation</div>
      <div class="learning-issue">1-by-1 tickets per learner for the same tech issue created backlog and extra admin effort on SAP's side.</div>
      <div class="learning-action">Partner consolidates tickets by issue type before raising to SAP.</div>
    </div>
  </div>
  <div class="learning-item">
    <div class="learning-num">4</div>
    <div class="learning-body">
      <div class="learning-title">Learner Access Ownership</div>
      <div class="learning-issue">Access queries (wrong user ID, inactive license, PRM status) reached SAP but could have been resolved on the partner side.</div>
      <div class="learning-action">Partner verifies and resolves access issues before escalating to SAP.</div>
    </div>
  </div>
  <div class="learning-item">
    <div class="learning-num">5</div>
    <div class="learning-body">
      <div class="learning-title">Pre-Exam Readiness</div>
      <div class="learning-issue">Local device and network issues blocked exam access and generated avoidable support cases.</div>
      <div class="learning-action">Require learners to complete the Yoodli technical readiness guide before attempting the exam.</div>
    </div>
  </div>
</div>

<div class="footer">GTS Management &nbsp;.&nbsp; Accenture x SAP Certification Week &nbsp;.&nbsp; $SprintLabel &nbsp;.&nbsp; As of $ReportDate</div>

</body>
</html>
"@

$html | Out-File $outputFile -Encoding UTF8
Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $outputFile" -ForegroundColor Yellow
Start-Process $outputFile
