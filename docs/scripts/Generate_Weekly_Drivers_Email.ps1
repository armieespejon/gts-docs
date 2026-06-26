# GTS June Weekly Drivers Email Generator

$To      = "your.recipient@sap.com"   # update before sending
$CC      = ""
$Today   = [datetime]::Today.ToString("MMMM d, yyyy")
$Subject = "GTS ESM | June Case Volume & Issue Drivers | $Today"

$body  = "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1a1a2e;max-width:760px'>"
$body += "<div style='background:#1e3a5f;padding:16px 20px;border-radius:6px 6px 0 0'>"
$body += "<div style='font-size:18px;font-weight:800;color:#fff'>GTS ESM | June Case Volume &amp; Issue Drivers</div>"
$body += "<div style='font-size:13px;color:#93c5fd;margin-top:4px'>Weekly Analysis &middot; $Today</div>"
$body += "</div>"

$body += "<div style='background:#f8fafc;padding:16px 20px;border:1px solid #e5e7eb;border-top:none'>"
$body += "<p style='margin:0 0 6px;font-size:13px'>Hi team,</p>"
$body += "<p style='margin:0 0 14px;font-size:13px'>Here is the June weekly case volume and issue driver breakdown.</p>"

# Key Takeaways
$body += "<div style='background:#fffbeb;border-left:4px solid #F0AB00;border-radius:4px;padding:12px 16px;margin-bottom:20px;font-size:13px'>"
$body += "<div style='font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;color:#92400e;margin-bottom:8px'>&#128161; Key Takeaways</div>"
$body += "<ul style='margin:0;padding-left:18px;line-height:1.8'>"
$body += "<li><strong>June running total: 8,072 cases</strong> &mdash; SBA: 6,484 (80%) | Practice Systems: 1,452 (18%) | Other: 136 (2%).</li>"
$body += "<li>Post-ACN sprint, volume has <strong>stabilized at 235&ndash;277 cases/day</strong> in W2&ndash;W4 &mdash; driven by normal SBA exam activity.</li>"
$body += "<li><strong>C_CPI is the top concern</strong> &mdash; 503 errors and slot unavailability are persisting into W4 (161 cases this week).</li>"
$body += "<li><strong>THR81 / C_THR81 is rising</strong> &mdash; 110 cases in W4 (90 SBA + 20 Practice Systems), up from 38 in W2. C_THR81 SBA exam issues are the main driver; 20 PS cases are residual from the Jun 19&ndash;22 maintenance window.</li>"
$body += "<li><strong>Exam results queries trending up</strong> &mdash; 120 in W4 vs 69 in W2 as learners follow up post-exam cycle.</li>"
$body += "<li>C_TS452/C_TS462 W3 spike (247 cases) was tied to the Jun 19&ndash;20 system unavailability.</li>"
$body += "</ul>"
$body += "</div>"

# Weekly Volume Table
$body += "<div style='font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;color:#1e3a5f;margin-bottom:8px'>Weekly Volume</div>"
$body += "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;font-size:13px;margin-bottom:20px'>"
$body += "<thead><tr style='background:#1e3a5f'>"
$body += "<th style='padding:8px 12px;text-align:left;color:#fff;font-weight:600'>Week</th>"
$body += "<th style='padding:8px 12px;text-align:center;color:#93c5fd;font-weight:600'>Total Cases</th>"
$body += "<th style='padding:8px 12px;text-align:center;color:#fde68a;font-weight:600'>Avg/Day</th>"
$body += "<th style='padding:8px 12px;text-align:center;color:#6ee7b7;font-weight:600'>Closed</th>"
$body += "<th style='padding:8px 12px;text-align:left;color:#bfdbfe;font-weight:600'>Note</th>"
$body += "</tr></thead><tbody>"
$body += "<tr style='border-bottom:1px solid #f3f4f6;background:#fff7ed'><td style='padding:8px 12px;font-weight:600;color:#92400e'>W1 Jun 1&ndash;7</td><td style='padding:8px 12px;text-align:center;font-weight:700;color:#92400e'>3,117</td><td style='padding:8px 12px;text-align:center;font-weight:700;color:#92400e'>445</td><td style='padding:8px 12px;text-align:center'>2,868</td><td style='padding:8px 12px;color:#92400e;font-weight:600'>Accenture AI Cert Week spike</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>W2 Jun 8&ndash;14</td><td style='padding:8px 12px;text-align:center'>1,644</td><td style='padding:8px 12px;text-align:center'>235</td><td style='padding:8px 12px;text-align:center'>1,321</td><td style='padding:8px 12px;color:#6b7280'>CPI + AIG tail-end. High slot &amp; access issues.</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>W3 Jun 15&ndash;21</td><td style='padding:8px 12px;text-align:center'>1,938</td><td style='padding:8px 12px;text-align:center'>277</td><td style='padding:8px 12px;text-align:center'>1,380</td><td style='padding:8px 12px;color:#6b7280'>TS452/TS462 spike. System unavailability Jun 19&ndash;20. THR81 &amp; C_THR81 maintenance impact.</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>W4 Jun 22&ndash;26</td><td style='padding:8px 12px;text-align:center'>1,373</td><td style='padding:8px 12px;text-align:center'>275</td><td style='padding:8px 12px;text-align:center'>897</td><td style='padding:8px 12px;color:#6b7280'>CPI dominates (503/slot issues). THR81 rising. Exam results queries up.</td></tr>"
$body += "<tr style='background:#f1f5f9;font-weight:700;border-top:2px solid #e5e7eb'><td style='padding:8px 12px;color:#1e3a5f'>June Total</td><td style='padding:8px 12px;text-align:center;color:#1d4ed8'>8,072</td><td style='padding:8px 12px;text-align:center;color:#1d4ed8'>310</td><td style='padding:8px 12px;text-align:center;color:#166534'>6,466</td><td style='padding:8px 12px'></td></tr>"
$body += "</tbody></table>"

# Issue Drivers Table
$body += "<div style='font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;color:#1e3a5f;margin-bottom:8px'>Issue Drivers by Week (Post-ACN)</div>"
$body += "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;font-size:13px;margin-bottom:20px'>"
$body += "<thead><tr style='background:#1e3a5f'>"
$body += "<th style='padding:8px 12px;text-align:left;color:#fff;font-weight:600'>Driver / Course</th>"
$body += "<th style='padding:8px 12px;text-align:center;color:#93c5fd;font-weight:600'>W2</th>"
$body += "<th style='padding:8px 12px;text-align:center;color:#fde68a;font-weight:600'>W3</th>"
$body += "<th style='padding:8px 12px;text-align:center;color:#6ee7b7;font-weight:600'>W4</th>"
$body += "<th style='padding:8px 12px;text-align:left;color:#bfdbfe;font-weight:600'>Trend</th>"
$body += "</tr></thead><tbody>"
$body += "<tr style='border-bottom:1px solid #f3f4f6;background:#eff6ff'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>Exam &ndash; SBA (L1)</td><td style='padding:8px 12px;text-align:center'>1,203 (73%)</td><td style='padding:8px 12px;text-align:center'>1,554 (80%)</td><td style='padding:8px 12px;text-align:center'>1,126 (82%)</td><td style='padding:8px 12px;color:#166534;font-weight:600'>&#8593; Growing share &mdash; normal cert activity</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>C_CPI</td><td style='padding:8px 12px;text-align:center'>172</td><td style='padding:8px 12px;text-align:center'>131</td><td style='padding:8px 12px;text-align:center;font-weight:700;color:#991b1b'>161 &#9888;</td><td style='padding:8px 12px;color:#991b1b'>503 errors and slot unavailability persisting W4</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>THR81</td><td style='padding:8px 12px;text-align:center'>38</td><td style='padding:8px 12px;text-align:center'>50</td><td style='padding:8px 12px;text-align:center;font-weight:700;color:#854d0e'>82 &#8593;</td><td style='padding:8px 12px;color:#854d0e'>Rising &mdash; residual maintenance impact</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>C_TS452 / C_TS462</td><td style='padding:8px 12px;text-align:center'>81</td><td style='padding:8px 12px;text-align:center;font-weight:700;color:#991b1b'>247 &#9888;</td><td style='padding:8px 12px;text-align:center'>81</td><td style='padding:8px 12px;color:#6b7280'>W3 spike tied to Jun 19&ndash;20 system unavailability.</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>C_ABAPD</td><td style='padding:8px 12px;text-align:center'>94</td><td style='padding:8px 12px;text-align:center'>48</td><td style='padding:8px 12px;text-align:center'>28</td><td style='padding:8px 12px;color:#166534'>&#8595; Declining week over week</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>Exam Results queries</td><td style='padding:8px 12px;text-align:center'>69</td><td style='padding:8px 12px;text-align:center'>87</td><td style='padding:8px 12px;text-align:center;font-weight:700;color:#854d0e'>120 &#8593;</td><td style='padding:8px 12px;color:#854d0e'>Growing &mdash; learners following up post-exam cycle</td></tr>"
$body += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-weight:600;color:#1e3a5f'>Slot unavailability</td><td style='padding:8px 12px;text-align:center'>118</td><td style='padding:8px 12px;text-align:center'>63</td><td style='padding:8px 12px;text-align:center'>58</td><td style='padding:8px 12px;color:#166534'>&#8595; Easing since W2 peak</td></tr>"
$body += "</tbody></table>"

$body += "<p style='margin:0;font-size:13px;color:#374151'>Happy to deep dive on any of these further. Let me know!</p>"
$body += "</div>"
$body += "<div style='background:#f1f5f9;padding:10px 20px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px;font-size:13px;color:#9ca3af'>"
$body += "GTS ESM Operations &middot; June Weekly Drivers Analysis &middot; $Today"
$body += "</div></body></html>"

# Generate EML
$emlPath   = "$env:TEMP\GTS_June_Weekly_Drivers_$([datetime]::Today.ToString('MMMdd')).eml"
$boundary  = "----=_GTS_$(Get-Random)"
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
$bodyB64   = [Convert]::ToBase64String($bodyBytes)
$ccField   = if ($CC) { "CC: $CC`r`n" } else { "" }
$eml  = "From: GTS ESM Operations`r`n"
$eml += "To: $To`r`n"
$eml += $ccField
$eml += "Subject: $Subject`r`n"
$eml += "MIME-Version: 1.0`r`n"
$eml += "Content-Type: multipart/mixed; boundary=`"$boundary`"`r`n`r`n"
$eml += "--$boundary`r`nContent-Type: text/html; charset=UTF-8`r`nContent-Transfer-Encoding: base64`r`n`r`n$bodyB64`r`n"
$eml += "--$boundary--`r`n"
[System.IO.File]::WriteAllText($emlPath, $eml, [System.Text.Encoding]::UTF8)
Write-Host "Done! Saved to: $emlPath"
Start-Process $emlPath
