param(
    [string]$Date         = (Get-Date -Format "yyyy-MM-dd"),
    [string]$OutputFolder = "C:\RejectionReports\output"
)

$SERVER     = "localhost"
$ADMIN_DB   = "XDSDATAADMIN"
$STAGING_DB = "XDSDATASTAGING"

# FormatID per table prefix - used to filter ValidationRules correctly
$FORMAT_IDS = @{
    "CDRN" = 7
    "CONJ" = 30
    "COMJ" = 31
}

$LogFolder = Join-Path $OutputFolder "logs"
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
if (-not (Test-Path $LogFolder))    { New-Item -ItemType Directory -Path $LogFolder    -Force | Out-Null }
$LogFile = Join-Path $LogFolder ("run_" + (Get-Date -Format "yyyyMMdd") + ".log")

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "{0}  {1,-8}  {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

function Invoke-SQL {
    param([string]$Database, [string]$Query)
    $connStr = "Server=$SERVER;Database=$Database;Integrated Security=True;"
    $conn    = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $cmd     = New-Object System.Data.SqlClient.SqlCommand($Query, $conn)
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $table   = New-Object System.Data.DataTable
    try {
        $conn.Open()
        $adapter.Fill($table) | Out-Null
    }
    finally {
        $conn.Close()
    }
    return ,$table
}

function Get-StagingTables {
    param([string]$DatePattern)
    Write-Log "Searching filelog for files created on: $DatePattern"
    $query  = "SELECT StagingName, FileName FROM filelog WHERE CAST(Createdondate AS DATE) = '$DatePattern' AND StagingName IS NOT NULL AND StagingName <> '' ORDER BY StagingName"
    $result = Invoke-SQL -Database $ADMIN_DB -Query $query
    $fileMap = @{}
    foreach ($row in $result.Rows) {
        $staging = [string]$row["StagingName"]
        $fname   = [string]$row["FileName"]
        $fname   = [System.IO.Path]::GetFileNameWithoutExtension($fname)
        if ($staging -match "^(COMJ|CONJ|CDRN)_") {
            $fileMap[$staging] = $fname
        }
    }
    return $fileMap
}

function Get-Rejections {
    param([string]$StagingTable)
    Write-Log "Pulling rejections from: $StagingTable"
    if ($StagingTable -match "^CDRN_") { $rejectCol = "IsRejected" } else { $rejectCol = "Rejected" }
    $query = "SELECT * FROM [$StagingTable] WHERE $rejectCol = 1"
    [System.Data.DataTable]$dt = Invoke-SQL -Database $STAGING_DB -Query $query
    return ,$dt
}

function Get-ValidationRules {
    param([int]$FormatID)
    Write-Log "Loading ValidationRules for FormatID=$FormatID from XDSDATAADMIN"
    $query  = "SELECT VRuleRef, VRuleDescription, VRuleCategory, VRuleType, VRejectionType FROM ValidationRules WHERE Active = 1 AND VFormatID = $FormatID"
    $table  = Invoke-SQL -Database $ADMIN_DB -Query $query
    $lookup = @{}
    if ($table -ne $null -and $table.Rows.Count -gt 0) {
        foreach ($row in $table.Rows) {
            $ref = [string]$row["VRuleRef"]
            if ($ref -ne $null -and $ref.Trim() -ne "") {
                # Strip "Rule " prefix to get the number e.g. "Rule 15" -> "15"
                $key = $ref.Trim() -replace "^Rule\s*", ""
                $lookup[$key] = @{
                    VRuleRef         = $ref.Trim()
                    VRuleDescription = [string]$row["VRuleDescription"]
                    VRuleCategory    = [string]$row["VRuleCategory"]
                    VRuleType        = [string]$row["VRuleType"]
                    VRejectionType   = [string]$row["VRejectionType"]
                }
            }
        }
    }
    Write-Log "Loaded $($lookup.Count) rules into lookup"
    return $lookup
}

function Get-EnrichedRejections {
    param([object]$Rejections, [hashtable]$RulesLookup)
    [System.Data.DataTable]$dt2 = $Rejections
    Write-Log "Enriching $($dt2.Rows.Count) rejection records"
    $allColumns = $dt2.Columns | Select-Object -ExpandProperty ColumnName
    $enriched = New-Object System.Collections.Generic.List[PSObject]
    foreach ($row in $dt2.Rows) {
        $rawRules = [string]$row["RejectedonRules"]
        $ruleIDs = ($rawRules -split "[|,]") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique
        foreach ($ruleID in $ruleIDs) {
            $rule = $RulesLookup[$ruleID]
            $obj = [ordered]@{
                RuleID          = $ruleID
                RuleRef         = if ($rule) { $rule.VRuleRef }         else { "Unknown" }
                RuleDescription = if ($rule) { $rule.VRuleDescription } else { "Unknown" }
                RuleCategory    = if ($rule) { $rule.VRuleCategory }    else { "" }
                RejectionType   = if ($rule) { $rule.VRejectionType }   else { "" }
                RuleType        = if ($rule) { $rule.VRuleType }        else { "" }
            }
            foreach ($col in $allColumns) {
                if ($col -ne "RejectedonRules") {
                    $obj[$col] = [string]$row[$col]
                }
            }
            $enriched.Add([PSCustomObject]$obj)
        }
    }
    Write-Log "Enriched to $($enriched.Count) rows"
    return $enriched
}

function Export-ToExcel {
    param(
        [System.Collections.Generic.List[PSObject]]$Data,
        [string]$StagingTable,
        [string]$FileName,
        [string]$OutFolder
    )
    $safeName = $FileName -replace '[\\/:*?"<>|]', '_'
    $filePath = Join-Path $OutFolder ($safeName + "_Rejections.xlsx")
    Write-Log "Writing Excel report: $filePath"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Add()
        $sheet    = $workbook.Worksheets.Item(1)
        $sheet.Name = "Rejections"

        # -- Row 1: Section title -----------------------------------------
        $titleCell = $sheet.Cells.Item(1, 1)
        $titleCell.Value2         = "REJECTION RULES IN THIS FILE"
        $titleCell.Font.Bold      = $true
        $titleCell.Font.Size      = 11
        $titleCell.Font.Color     = 16777215
        $titleCell.Interior.Color = 8015994

        # -- Rows 2+: One row per unique rule ----------------------------
        $uniqueRules = $Data | Group-Object RuleID | ForEach-Object { $_.Group[0] } | Sort-Object RuleRef
        $ruleRowIdx  = 2
        foreach ($rule in $uniqueRules) {
            $ruleCell = $sheet.Cells.Item($ruleRowIdx, 1)
            $ruleCell.Value2         = "$($rule.RuleRef) - $($rule.RuleDescription)"
            $ruleCell.Font.Bold      = $false
            $ruleCell.Font.Italic    = $true
            $ruleCell.Font.Color     = 2629176
            $ruleCell.Interior.Color = 15132390
            $ruleRowIdx++
        }

        # -- Blank separator row ------------------------------------------
        $ruleRowIdx++

        # -- Column headers -----------------------------------------------
        $headers   = $Data[0].PSObject.Properties.Name
        $headerRow = $ruleRowIdx

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $sheet.Cells.Item($headerRow, $c + 1)
            $cell.Value2              = $headers[$c]
            $cell.Font.Bold           = $true
            $cell.Font.Color          = 16777215
            $cell.Interior.Color      = 8015994
            $cell.HorizontalAlignment = -4108
        }
        $sheet.Rows.Item($headerRow).RowHeight = 20

        # -- Data rows ----------------------------------------------------
        $rowIdx = $headerRow + 1
        foreach ($item in $Data) {
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $sheet.Cells.Item($rowIdx, $c + 1).Value2 = $item.($headers[$c])
            }
            if ($rowIdx % 2 -eq 0) {
                $sheet.Rows.Item($rowIdx).Interior.Color = 15921906
            }
            $rowIdx++
        }

        $sheet.UsedRange.Columns.AutoFit() | Out-Null
        $sheet.Application.ActiveWindow.SplitRow    = $headerRow
        $sheet.Application.ActiveWindow.FreezePanes = $true

        # -- Summary sheet ------------------------------------------------
        $summary = $workbook.Worksheets.Add()
        $summary.Name = "Summary"
        $summaryRows = @(
            @("File name",        $FileName),
            @("Staging table",    $StagingTable),
            @("Total rejections", $Data.Count),
            @("Unique rules",     ($Data | Select-Object -ExpandProperty RuleID -Unique).Count),
            @("Extracted on",     (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
        )
        for ($i = 0; $i -lt $summaryRows.Count; $i++) {
            $summary.Cells.Item($i+1, 1).Value2    = $summaryRows[$i][0]
            $summary.Cells.Item($i+1, 1).Font.Bold = $true
            $summary.Cells.Item($i+1, 2).Value2    = $summaryRows[$i][1]
        }
        $summary.Columns.AutoFit() | Out-Null
        $sheet.Activate()
        $workbook.SaveAs($filePath, 51)
        Write-Log "Saved: $filePath"
        return $filePath
    }
    finally {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# -- Main ---------------------------------------------------------------------
Write-Log "============================================================"
Write-Log "Rejection Extraction Pipeline - starting"
Write-Log "Date        : $Date"
Write-Log "Output      : $OutputFolder"
Write-Log "============================================================"

try {
    $tables = Get-StagingTables -DatePattern $Date
} catch {
    Write-Log "ERROR querying filelog: $_" "ERROR"
    exit 1
}

if (-not $tables -or $tables.Count -eq 0) {
    Write-Log "No staging tables found for date: $Date" "WARN"
    Write-Log "Nothing to process - exiting."
    exit 0
}

$tableList = $tables.Keys
Write-Log "Found $($tableList.Count) staging table(s): $($tableList -join ', ')"

$successCount = 0
$failCount    = 0

foreach ($table in $tableList) {
    $friendlyName = $tables[$table]
    Write-Log "------------------------------------------------------------"
    Write-Log "File: $friendlyName"
    try {
        # Get FormatID for this table prefix
        $prefix   = ($table -split "_")[0]
        $formatID = $FORMAT_IDS[$prefix]
        Write-Log "Using FormatID=$formatID for prefix $prefix"

        $rulesLookup = Get-ValidationRules -FormatID $formatID
        Write-Log "Loaded $($rulesLookup.Count) rules for this file type"

        $rejections = Get-Rejections -StagingTable $table
        if ($rejections -eq $null -or $rejections.Rows.Count -eq 0) {
            Write-Log "No rejections found in $table - skipping" "WARN"
            continue
        }
        Write-Log "Found $($rejections.Rows.Count) rejected records"
        $enriched = Get-EnrichedRejections -Rejections $rejections -RulesLookup $rulesLookup
        $path     = Export-ToExcel -Data $enriched -StagingTable $table -FileName $friendlyName -OutFolder $OutputFolder
        $successCount++
    } catch {
        Write-Log "ERROR processing $table : $_" "ERROR"
        $failCount++
    }
}

Write-Log "============================================================"
Write-Log "Run complete - $successCount succeeded, $failCount failed"
Write-Log "============================================================"

if ($failCount -gt 0) { exit 1 }
