[setup.md](https://github.com/user-attachments/files/29468144/setup.md)
# Setup Guide

## Step 1 -- Verify prerequisites

Open PowerShell and check your version:

```powershell
$PSVersionTable.PSVersion
```

You need Major version 4 or higher.

Verify Excel is installed:

```powershell
New-Object -ComObject Excel.Application
```

If this errors, Excel is not installed on the server.

## Step 2 -- Test your SQL Server connection

Run this in PowerShell to confirm connectivity:

```powershell
$conn = New-Object System.Data.SqlClient.SqlConnection(
    "Server=YOUR_SERVER;Database=XDSDATAADMIN;Integrated Security=True;"
)
$conn.Open()
Write-Host "Connection state: $($conn.State)"
$conn.Close()
```

Expected output: `Connection state: Open`

## Step 3 -- Verify your FormatIDs

Run this in SSMS to confirm the FormatID for each of your file types:

```sql
SELECT DISTINCT FormatID, FileName
FROM XDSDATAADMIN..filelog
WHERE StagingName LIKE 'CDRN%'
   OR StagingName LIKE 'COMJ%'
   OR StagingName LIKE 'CONJ%'
ORDER BY FormatID
```

Update the `$FORMAT_IDS` hashtable in the script to match.

## Step 4 -- Create your output folder

The script creates this automatically, but you can pre-create it:

```powershell
New-Item -ItemType Directory -Path "C:\RejectionReports\output" -Force
New-Item -ItemType Directory -Path "C:\RejectionReports\output\logs" -Force
```

## Step 5 -- Run a test

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\RejectionPipeline.ps1
```

Check the log file to confirm everything ran correctly.

## Step 6 -- Schedule with Task Scheduler

See the README for full Task Scheduler setup instructions.

## Troubleshooting

**"Cannot open database requested by the login"**
Your Windows account does not have access to that database. Contact your DBA.

**"No staging tables found for date"**
No files were loaded today matching the CDRN/COMJ/CONJ prefixes. 
Try running for a date you know had files: `.\RejectionPipeline.ps1 -Date "2026-06-14"`

**"Property RowID cannot be found"**
The staging table for this file type does not have a RowID column. 
The script handles this automatically with dynamic column detection.

**Excel files are created but rule descriptions show Unknown**
Check that your `$FORMAT_IDS` values match the FormatID in filelog for each table prefix.
Run the verification query in Step 3 above.
