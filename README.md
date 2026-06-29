# SQL Server Rejection Extraction Pipeline

> Automated daily pipeline that extracts rejected credit bureau records from SQL Server, enriches them with validation rule descriptions, and delivers formatted Excel reports — replacing a manual 20-minute process with a scheduled script that runs in under 60 seconds.

---

## Business Problem

In a credit bureau environment, data files submitted by member organisations go through a validation process before being loaded. Records that fail validation are flagged as **rejected** and stored in staging tables. 

Previously, extracting these rejections required a data operator to:
1. Manually query the filelog to find staging table names
2. Run rejection queries against each staging table
3. Cross-reference rule IDs against a validation rules database
4. Copy results into a formatted Excel file for distribution

This process took approximately 20 minutes each morning and was prone to human error.

---

## Solution

A PowerShell pipeline that automates the entire workflow end-to-end, scheduled via Windows Task Scheduler to run at 08:30 every morning.

```
filelog (XDSDATAADMIN)
        |
        | -- finds staging tables created today
        v
XDSDATASTAGING..[StagingTable]
        |
        | -- extracts WHERE IsRejected = 1 / Rejected = 1
        v
ValidationRules (XDSDATAADMIN)
        |
        | -- enriches each rejection with rule description
        | -- filtered by VFormatID (file type specific)
        v
Formatted Excel Report
        |
        | -- rule summary header block
        | -- alternating row colours
        | -- frozen header row
        | -- summary sheet
        v
Output folder (auto-created)
```

---

## Tech Stack

| Component | Technology |
|---|---|
| Language | PowerShell 4+ |
| Database | SQL Server 2012 |
| Output | Excel (via COM automation) |
| Scheduling | Windows Task Scheduler |
| Authentication | Windows Authentication (Integrated Security) |

---

## Features

- **Zero manual input** -- runs automatically every morning
- **Dynamic table detection** -- queries filelog by date to find relevant staging tables
- **Format-aware rule lookup** -- filters ValidationRules by VFormatID so each file type gets the correct rule descriptions
- **Multi-table support** -- processes CDRN, COMJ and CONJ file types in a single run
- **Rule summary header** -- each Excel file opens with a clear list of all rules violated in that file
- **Formatted output** -- dark blue headers, alternating row fills, frozen panes, auto-fitted columns
- **Daily log file** -- every run writes a timestamped log so you can audit what ran and when
- **Manual override** -- can be run for any specific date via command-line parameter

---

## Project Structure

```
sql-to-excel-rejection-pipeline/
|-- RejectionPipeline.ps1      # Main script
|-- README.md                  # This file
|-- docs/
|   |-- architecture.png       # Pipeline flow diagram
|   |-- setup.md               # Setup and configuration guide
|-- sample_output/
    |-- sample_report.md       # Description of Excel output format
```

---

## Setup

### Prerequisites

- Windows Server with PowerShell 4 or higher
- SQL Server 2012 or later
- Microsoft Excel installed on the server
- Windows account with read access to:
  - `XDSDATAADMIN..filelog`
  - `XDSDATAADMIN..ValidationRules`
  - `XDSDATASTAGING..[staging tables]`

### Configuration

Open `RejectionPipeline.ps1` and update these variables at the top of the script:

```powershell
# Server name
$SERVER = "localhost"   # or your SQL Server instance name

# Output folder - will be created automatically if it does not exist
$OutputFolder = "C:\RejectionReports\output"

# FormatID per table prefix - update if your format IDs differ
$FORMAT_IDS = @{
    "CDRN" = 7
    "CONJ" = 30
    "COMJ" = 31
}
```

### Running manually

```powershell
# Run for today
.\RejectionPipeline.ps1

# Run for a specific date
.\RejectionPipeline.ps1 -Date "2026-06-14"

# Run with a custom output folder
.\RejectionPipeline.ps1 -OutputFolder "D:\Reports"
```

### Scheduling with Windows Task Scheduler

1. Open Task Scheduler
2. Create Task (not Basic Task)
3. General tab: tick "Run whether user is logged on or not" and "Run with highest privileges"
4. Triggers tab: Daily at 08:30
5. Actions tab:
   - Program: `powershell.exe`
   - Arguments: `-ExecutionPolicy Bypass -File "C:\Scripts\RejectionPipeline.ps1"`

---

## Excel Output Format

Each report is named after the original source file:

```
L@W JUDGMENTS - 2026-06-14T01.13.01(Commercial)_Rejections.xlsx
FlaggedStatusCustomers-13062026_Rejections.xlsx
```

The Rejections sheet is structured as:

| Row | Content |
|---|---|
| 1 | "REJECTION RULES IN THIS FILE" (dark blue header) |
| 2..n | One row per unique rule: `Rule 15 - ID Number or Passport Number not provided` |
| n+1 | Blank separator |
| n+2 | Column headers (dark blue, white text, frozen) |
| n+3+ | Rejection data rows (alternating fill) |

A second **Summary** sheet shows file name, staging table, total rejections, unique rules, and extraction timestamp.

---

## Log Files

Each run writes a log to `[OutputFolder]\logs\run_YYYYMMDD.log`:

```
2026-06-14 08:30:01  INFO      ============================================================
2026-06-14 08:30:01  INFO      Rejection Extraction Pipeline - starting
2026-06-14 08:30:01  INFO      Date        : 2026-06-14
2026-06-14 08:30:02  INFO      Found 3 staging table(s): CDRN_..., COMJ_..., CONJ_...
2026-06-14 08:30:02  INFO      Loaded 245 rules for this file type
2026-06-14 08:30:03  INFO      Found 4 rejected records
2026-06-14 08:30:05  INFO      Saved: C:\RejectionReports\output\FlaggedStatusCustomers_Rejections.xlsx
2026-06-14 08:30:05  INFO      Run complete - 3 succeeded, 0 failed
```

---

## Background

This pipeline was built as part of a data automation initiative at a South African credit bureau, where the author works as a Senior Data Operator with six years of experience in SQL Server, SSIS, and data pipeline operations.

The project demonstrates practical data engineering skills applied to a real production environment:
- ETL pipeline design and implementation
- Multi-database querying with Windows Authentication
- Data enrichment and transformation in-memory
- Automated reporting with formatted Excel output
- Production scheduling and logging

---

## Author

**Nonhlanhla** | Senior Data Operator transitioning to Data Engineering

- 6 years experience with SQL Server, SSIS, and log shipping in a production credit bureau environment
- AWS Cloud Practitioner certified
- Azure Data Fundamentals certified
- Building toward: Python, Apache Airflow, dbt Core, Snowflake

---

## Licence

MIT -- free to use, adapt, and build on.
