# ============================================================
# GTS Coaching Log - Word Sample (HTML saved as .doc)
# Word opens this natively - no COM object needed
# ============================================================

$OutputFile = "C:\Users\I535893\Downloads\Claude and Me\GTS_CoachingLog_SAMPLE_Marviely.doc"
$AgentName  = "Marviely Quinto"
$Team       = "PH Team"
$ReportDate = "June 17, 2026"

$html = @"
<html xmlns:o='urn:schemas-microsoft-com:office:office'
      xmlns:w='urn:schemas-microsoft-com:office:word'
      xmlns='http://www.w3.org/TR/REC-html40'>
<head>
<meta charset='UTF-8'/>
<xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml>
<style>
  @page { size: 21cm 29.7cm; margin: 1.8cm 2cm; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 11pt; color: #1e293b; }

  /* TITLE BLOCK */
  .doc-title   { font-size: 22pt; font-weight: bold; color: #1e3a5f; margin: 0 0 4px 0; }
  .doc-agent   { font-size: 14pt; color: #007DB8; margin: 0 0 3px 0; }
  .doc-meta    { font-size: 9pt; color: #64748b; margin: 0 0 10px 0; }
  .divider     { border: none; border-top: 2px solid #e2e8f0; margin: 8px 0 12px 0; }

  /* SECTION LABELS */
  .section-lbl {
    font-size: 8pt; font-weight: bold; text-transform: uppercase;
    letter-spacing: 1px; color: #64748b;
    border-bottom: 1px solid #e2e8f0;
    margin: 16px 0 8px 0; padding-bottom: 4px;
  }

  /* KPI TABLE */
  .kpi-table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
  .kpi-table th {
    background: #1e3a5f; color: #fff;
    font-size: 8pt; font-weight: bold; text-align: center;
    padding: 6px 4px; border: 1px solid #fff;
  }
  .kpi-table td {
    background: #f8fafc; font-size: 20pt; font-weight: bold;
    text-align: center; padding: 8px 4px;
    border: 1px solid #e2e8f0;
  }
  .c-navy  { color: #1e3a5f; }
  .c-blue  { color: #007DB8; }
  .c-amber { color: #b45309; }
  .c-red   { color: #b91c1c; }
  .c-green { color: #047857; }
  .c-gray  { color: #64748b; }

  /* ENTRY BLOCK */
  .entry { margin-bottom: 14px; border: 1px solid #e2e8f0; border-radius: 4px; overflow: hidden; }
  .entry-header { background: #f1f5f9; }
  .entry-header table { width: 100%; border-collapse: collapse; }
  .entry-header td { padding: 7px 10px; font-size: 10pt; border: none; }
  .entry-date   { font-weight: bold; color: #007DB8; width: 30%; }
  .entry-topic  { font-weight: bold; color: #1e3a5f; text-align: center; width: 45%; }
  .entry-status { font-weight: bold; text-align: right; width: 25%; }
  .status-open     { color: #b91c1c; }
  .status-closed   { color: #047857; }
  .status-progress { color: #b45309; }

  .entry-body { padding: 10px 12px; }
  .entry-snapshot { font-size: 8.5pt; color: #64748b; margin-bottom: 8px; }
  .entry-lbl  { font-size: 9pt; font-weight: bold; color: #1e3a5f; margin: 8px 0 3px 0; }
  .entry-text { font-size: 10pt; color: #374151; margin: 0 0 8px 0; line-height: 1.5; }
  .action-box {
    background: #f8fafc; border-left: 3px solid #b45309;
    padding: 8px 12px; margin-top: 6px;
    font-size: 10pt; color: #1e293b; line-height: 1.8;
  }
  .action-lbl { font-weight: bold; color: #b45309; font-size: 8.5pt; text-transform: uppercase; }

  /* NEW ENTRY TEMPLATE */
  .new-entry-header { background: #f1f5f9; border: 1px solid #e2e8f0; }
  .new-entry-header table { width: 100%; border-collapse: collapse; }
  .new-entry-header td { padding: 7px 10px; font-size: 10pt; color: #94a3b8; border: none; }
  .note-lbl   { font-size: 9pt; font-weight: bold; color: #1e3a5f; margin: 12px 0 4px 0; }
  .note-line  { border: none; border-bottom: 1px solid #cbd5e1; margin: 0 0 10px 0; height: 18px; display: block; }

  /* FOOTER */
  .footer { font-size: 8pt; color: #94a3b8; text-align: center; margin-top: 20px; border-top: 1px solid #e2e8f0; padding-top: 8px; }

  /* PAGE BREAK */
  .page-break { page-break-before: always; }
</style>
</head>
<body>

<!-- TITLE BLOCK -->
<p class='doc-title'>GTS ESM &mdash; Manager Coaching Log</p>
<p class='doc-agent'>$AgentName &nbsp;&nbsp;&mdash;&nbsp;&nbsp; $Team</p>
<p class='doc-meta'>Last updated: $ReportDate &nbsp;&nbsp;&mdash;&nbsp;&nbsp; GTS Management &nbsp;&nbsp;&mdash;&nbsp;&nbsp; Confidential</p>
<hr class='divider'/>

<!-- CURRENT PERFORMANCE SNAPSHOT -->
<div class='section-lbl'>Current Performance Snapshot &mdash; as of $ReportDate</div>
<table class='kpi-table'>
  <tr>
    <th>YTD Volume</th>
    <th>Active Cases</th>
    <th>Closure Rate</th>
    <th>CSAT</th>
    <th>Avg Age</th>
  </tr>
  <tr>
    <td class='c-navy'>312</td>
    <td class='c-amber'>28</td>
    <td class='c-red'>81%</td>
    <td class='c-red'>65%</td>
    <td class='c-amber'>18d</td>
  </tr>
</table>

<!-- ACTIVE CASES BY AGE -->
<div class='section-lbl'>Active Cases by Age</div>
<table class='kpi-table'>
  <tr>
    <th>0 &ndash; 7 Days</th>
    <th>8 &ndash; 15 Days</th>
    <th>16 &ndash; 30 Days</th>
    <th>30+ Days</th>
  </tr>
  <tr>
    <td class='c-navy'>12</td>
    <td class='c-blue'>8</td>
    <td class='c-amber'>5</td>
    <td class='c-red'>3</td>
  </tr>
</table>

<!-- COACHING LOG -->
<div class='section-lbl'>Coaching Log &mdash; Most Recent First</div>

<!-- ENTRY 1 -->
<div class='entry'>
  <div class='entry-header'>
    <table><tr>
      <td class='entry-date'>June 10, 2026</td>
      <td class='entry-topic'>Case Volume / Aging</td>
      <td class='entry-status status-open'>OPEN</td>
    </tr></table>
  </div>
  <div class='entry-body'>
    <div class='entry-snapshot'>Snapshot at session: Active 28 &nbsp;/&nbsp; 30+ days: 3 &nbsp;/&nbsp; Closure Rate: 81% &nbsp;/&nbsp; CSAT: 65%</div>
    <div class='entry-lbl'>Discussion</div>
    <p class='entry-text'>Discussed 3 cases aged beyond 30 days. Agent cited customer non-response as main blocker. Agreed to escalate 2 of the 3 cases to L2 by end of week. Agent to update case notes daily.</p>
    <div class='action-box'>
      <div class='action-lbl'>Agreed Action</div>
      Escalate 2 oldest cases to L2. Update all active case notes by Jun 13.<br/>
      <strong>Due Date:</strong> &nbsp;June 13, 2026
    </div>
  </div>
</div>

<!-- ENTRY 2 -->
<div class='entry'>
  <div class='entry-header'>
    <table><tr>
      <td class='entry-date'>June 3, 2026</td>
      <td class='entry-topic'>CSAT</td>
      <td class='entry-status status-closed'>CLOSED</td>
    </tr></table>
  </div>
  <div class='entry-body'>
    <div class='entry-snapshot'>Snapshot at session: Active 25 &nbsp;/&nbsp; 30+ days: 1 &nbsp;/&nbsp; Closure Rate: 83% &nbsp;/&nbsp; CSAT: 65%</div>
    <div class='entry-lbl'>Discussion</div>
    <p class='entry-text'>Reviewed 3 low-rated survey responses. Common theme: delayed first reply. Agent acknowledged and committed to same-day acknowledgement on all new cases.</p>
    <div class='action-box'>
      <div class='action-lbl'>Agreed Action</div>
      Same-day acknowledgement on all new cases. Review own CSAT weekly.<br/>
      <strong>Due Date:</strong> &nbsp;June 10, 2026 &nbsp;&mdash;&nbsp; Closed Jun 10
    </div>
  </div>
</div>

<!-- NEW ENTRY TEMPLATE -->
<div class='section-lbl'>New Entry &mdash; [Next 1:1 Date]</div>
<div class='new-entry-header'>
  <table><tr>
    <td style='width:33%'>Date: ___________</td>
    <td style='width:44%;text-align:center'>Topic: ___________</td>
    <td style='width:23%;text-align:right'>Status: ___________</td>
  </tr></table>
</div>
<p style='font-size:8.5pt;color:#94a3b8;margin:6px 0 10px 0;'>Snapshot: [auto-filled on next script run]</p>

<div class='note-lbl'>Discussion</div>
<span class='note-line'></span>
<span class='note-line'></span>
<span class='note-line'></span>
<span class='note-line'></span>

<div class='note-lbl'>Agreed Action</div>
<span class='note-line'></span>
<span class='note-line'></span>

<div class='note-lbl'>Due Date</div>
<span class='note-line'></span>

<div class='note-lbl'>Follow-up / Next Steps</div>
<span class='note-line'></span>
<span class='note-line'></span>

<div class='footer'>GTS Management &nbsp;&mdash;&nbsp; Coaching Log &nbsp;&mdash;&nbsp; $AgentName &nbsp;&mdash;&nbsp; Confidential</div>

</body>
</html>
"@

$html | Out-File $OutputFile -Encoding UTF8

Write-Host ""
Write-Host "Done! Saved to:" -ForegroundColor Green
Write-Host "  $OutputFile" -ForegroundColor Yellow
Start-Process $OutputFile
