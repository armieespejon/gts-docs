# GTS ESM PowerShell Scripts

All scripts are saved in `Downloads\Claude and Me\` and `OneDrive - SAP SE\Projects\2026\Claude Scripts for Mie\`.

To run any script: `powershell -ExecutionPolicy Bypass -File "script_name.ps1"`

## Report Generators

| Script | Output | Notes |
|---|---|---|
| `Generate_Exec_Summary.ps1` | `GTS_Executive_Summary.html` | Auto-detects latest cases + surveys xlsx |
| `Generate_Daily_Inflow.ps1` | `GTS_Daily_Inflow_June.html` | Weekly + daily inflow & closures, toggle, rate% |
| `Generate_Agent_Backlog.ps1` | `GTS_Agent_Backlog.html` | 5-run trend, history in `backlog_history.json` |
| `Generate_CSAT_Bulletin.ps1` | `GTS_CSAT_Bulletin.html` | CSAT, DSAT, voice of customer |
| `Generate_Winners_Circle.ps1` | `GTS_Winners_Circle_Wxx.html` | Update CONFIG block before running |
| `Generate_CoachingLog_HTML.ps1` | `GTS_CoachingLog.html` | Reads from `GTS_CoachingLog_Entries.csv` |
| `Generate_CoachingLog_Word.ps1` | One `.doc` per agent | Individual coaching logs |

## Email & Communication

| Script | Output | Notes |
|---|---|---|
| `Generate_Inflow_Email.ps1` | `.eml` draft file | WIP — Outlook COM issue, revisit |

## Data Requirements

All scripts auto-detect the latest xlsx from `C:\Users\I535893\Downloads\`:

- **Cases:** `GTS ESM Cases_[date].xlsx` — include `COMPLETED_ON` + `CLOSED_ON` columns for closure tracking
- **Surveys:** `GTS ESM Surveys_[date].xlsx`

## Download Scripts

| Script | Download |
|---|---|
| Generate_Daily_Inflow.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_Daily_Inflow.ps1) |
| Generate_Exec_Summary.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_Exec_Summary.ps1) |
| Generate_Agent_Backlog.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_Agent_Backlog.ps1) |
| Generate_CSAT_Bulletin.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_CSAT_Bulletin.ps1) |
| Generate_Winners_Circle.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_Winners_Circle.ps1) |
| Generate_CoachingLog_HTML.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_CoachingLog_HTML.ps1) |
| Generate_CoachingLog_Word.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_CoachingLog_Word.ps1) |
| Generate_Inflow_Email.ps1 | [Download](https://github.com/armieespejon/gts-docs/raw/main/docs/scripts/Generate_Inflow_Email.ps1) |
