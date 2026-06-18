# How to Run Reports

Step-by-step guide to generating the GTS ESM HTML reports.

## Prerequisites

- Latest cases xlsx dropped in `Downloads` folder
- Latest surveys xlsx dropped in `Downloads` folder

## Scripts

| Report | Script | Output |
|---|---|---|
| Executive Summary | `Generate_Exec_Summary.ps1` | `GTS_Executive_Summary.html` |
| CSAT Bulletin | `Generate_CSAT_Bulletin.ps1` | `GTS_CSAT_Bulletin.html` |
| Daily Inflow | `Generate_Daily_Inflow.ps1` | `GTS_Daily_Inflow_June.html` |
| Agent Backlog | `Generate_Agent_Backlog.ps1` | `GTS_Agent_Backlog.html` |
| Coaching Log | `Generate_CoachingLog_HTML.ps1` | `GTS_CoachingLog.html` |
| Winners Circle | `Generate_Winners_Circle.ps1` | `GTS_Winners_Circle.html` |

## How to run

```powershell
powershell -ExecutionPolicy Bypass -File "Generate_Exec_Summary.ps1"
```

All scripts are in `C:\Users\I535893\Downloads\Claude and Me\`

Outputs are saved to `C:\Users\I535893\OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\`
