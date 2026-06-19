# Data Sources

## Cases File

- **Filename pattern:** `GTS ESM Cases_[MonthDay].xlsx`
- **Sheet:** SERVICE_CLOUD
- **Key fields:** DISPLAYID, MONTH, CREATED_ON_DATE, STATUS, L1_CATEGORY, EXT_COURSECODE, PROCESSOR_ID, PROCESSOR_NAME, EXT_MISCINFO, SUBJECT, COMPLETED_ON, CLOSED_ON

### Closure Date Fields
When exporting cases, make sure to include these two additional columns:

| Field | Description |
|---|---|
| `COMPLETED_ON` | Date the case was marked Completed — **primary closure date** |
| `CLOSED_ON` | Date the case was marked Closed — used as fallback if COMPLETED_ON is empty |

> **Important:** Always export with both fields. Scripts auto-detect their presence and will show the Closure row and Rate% in the Daily Inflow report only when these columns are available.

**Closure definition:** A case is considered closed when its STATUS is `Closed` OR `Completed`. The closure date used is `COMPLETED_ON` first, then `CLOSED_ON` as fallback.

## Surveys File

- **Filename pattern:** `GTS ESM Surveys_[MonthDay].xlsx`
- **Sheet:** Case List
- **Key fields:** MONTH, RATING, L1_CATEGORY, COMMENT1, COMMENT2

## Notes

- CSAT = Top-2-Box (Very Satisfied + Somewhat Satisfied)
- DE processors use `PROCESSOR_ID = SAP_GERMANY_USER`; resolve name from `EXT_MISCINFO` keyword
- ILT cases detected via `SUBJECT` or `EXT_MISCINFO` containing "ILT"
