# Data Sources

## Cases File

- **Filename pattern:** `GTS ESM Cases_[MonthDay].xlsx`
- **Sheet:** SERVICE_CLOUD
- **Key fields:** DISPLAYID, MONTH, CREATED_ON_DATE, STATUS, L1_CATEGORY, EXT_COURSECODE, PROCESSOR_ID, PROCESSOR_NAME, EXT_MISCINFO, SUBJECT

## Surveys File

- **Filename pattern:** `GTS ESM Surveys_[MonthDay].xlsx`
- **Sheet:** Case List
- **Key fields:** MONTH, RATING, L1_CATEGORY, COMMENT1, COMMENT2

## Notes

- CSAT = Top-2-Box (Very Satisfied + Somewhat Satisfied)
- DE processors use `PROCESSOR_ID = SAP_GERMANY_USER`; resolve name from `EXT_MISCINFO` keyword
- ILT cases detected via `SUBJECT` or `EXT_MISCINFO` containing "ILT"
