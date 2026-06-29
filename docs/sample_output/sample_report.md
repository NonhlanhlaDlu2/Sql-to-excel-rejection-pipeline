[sample_report.md](https://github.com/user-attachments/files/29468223/sample_report.md)
# Sample Output

## File naming

Reports are named after the original source file from filelog, not the staging table name:

```
FlaggedStatusCustomers-13062026_Rejections.xlsx
L_W_JUDGMENTS_-_2026-06-14T01_13_01_Commercial__Rejections.xlsx
L_W_JUDGMENTS_-_2026-06-14T01_10_01_Consumer__Rejections.xlsx
```

## Rejections sheet layout

```
+----------------------------------------------------------+
| REJECTION RULES IN THIS FILE                             |  <- dark blue, white bold
+----------------------------------------------------------+
| Rule 15 - ID Number or Passport Number not provided      |  <- italic, light blue fill
| Rule 5  - Gender field is blank or invalid               |  <- italic, light blue fill
+----------------------------------------------------------+
|          (blank separator row)                           |
+----------------------------------------------------------+
| RuleID | RuleRef | RuleDescription | RuleCategory | ...  |  <- dark blue headers, frozen
+----------------------------------------------------------+
| 15     | Rule 15 | ID Number or... | Demographic  | ...  |  <- alternating row fill
| 5      | Rule 5  | Gender field... | Demographic  | ...  |
+----------------------------------------------------------+
```

## Summary sheet

| Field | Value |
|---|---|
| File name | FlaggedStatusCustomers-13062026 |
| Staging table | CDRN_2147670_20260614064704546 |
| Total rejections | 4 |
| Unique rules | 2 |
| Extracted on | 2026-06-14 08:30:05 |

## Log file sample

```
2026-06-14 08:30:01  INFO      ============================================================
2026-06-14 08:30:01  INFO      Rejection Extraction Pipeline - starting
2026-06-14 08:30:01  INFO      Date        : 2026-06-14
2026-06-14 08:30:01  INFO      Output      : C:\RejectionReports\output
2026-06-14 08:30:01  INFO      ============================================================
2026-06-14 08:30:01  INFO      Searching filelog for files created on: 2026-06-14
2026-06-14 08:30:02  INFO      Found 3 staging table(s): CDRN_..., COMJ_..., CONJ_...
2026-06-14 08:30:02  INFO      ------------------------------------------------------------
2026-06-14 08:30:02  INFO      File: FlaggedStatusCustomers-13062026
2026-06-14 08:30:02  INFO      Using FormatID=7 for prefix CDRN
2026-06-14 08:30:02  INFO      Loaded 245 rules for this file type
2026-06-14 08:30:03  INFO      Found 4 rejected records
2026-06-14 08:30:03  INFO      Enriching 4 rejection records
2026-06-14 08:30:03  INFO      Enriched to 5 rows
2026-06-14 08:30:05  INFO      Saved: C:\RejectionReports\output\FlaggedStatusCustomers-13062026_Rejections.xlsx
2026-06-14 08:30:05  INFO      ============================================================
2026-06-14 08:30:05  INFO      Run complete - 3 succeeded, 0 failed
2026-06-14 08:30:05  INFO      ============================================================
```
