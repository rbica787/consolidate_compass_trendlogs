# ============================================================
# Consolidate Trendlog CSV Files by BACnet Device Instance
#
# Example input:
# Dev 10423, BI 6, present-value, UV-3, OCCUPANCY_SNSR-M-2026-09.csv
#
# Device instance = number immediately following "Dev"
#
# Output:
# Dev 10423 ALL TRENDLOGS 20260910_161100.xlsx
#
# Behavior:
# - Recursively searches all subfolders
# - Blank device selection = process ALL devices
# - Enter specific device instance = process only that device
# - Blank output location = use source directory as the base save path
# - Creates a "###ALL TRENDLOGS" subdirectory inside the selected/base save path
# - Creates ONE worksheet per device workbook, named with the unit label
# - Trend columns use BACnet point name from the filename plus point description
# - Before the final repeat prompt, optionally consolidates ONLY the workbooks
#   created during the current run into one master workbook
# - Master workbook preserves each source worksheet as a separate tab
# - Device instance is retained in each master workbook tab name
# - Consolidation routine supports CSV, XLS, XLSX, XLSM, and XLSB sources
# - Blank final prompt = exit
# - Any entry at final prompt = run again
# ============================================================

$ErrorActionPreference = "Stop"


# ============================================================
# MAIN REPEAT LOOP
# ============================================================

do {

    Clear-Host

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " Trendlog Consolidation by Device Instance" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""


    # ========================================================
    # SOURCE DIRECTORY
    # ========================================================

    do {

        $SourceDirectory = Read-Host "Enter the directory containing the trendlog CSV files"

        $SourceDirectory = $SourceDirectory.Trim().Trim('"')


        if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {

            Write-Host ""
            Write-Host "Directory does not exist. Please try again." -ForegroundColor Red
            Write-Host ""

            $SourceDirectory = $null
        }

    }
    until ($SourceDirectory)


    $SourceDirectory = (
        Resolve-Path -LiteralPath $SourceDirectory
    ).Path


    # ========================================================
    # OUTPUT DIRECTORY
    # ========================================================

    Write-Host ""
    Write-Host "Output location:" -ForegroundColor Yellow
    Write-Host "Press ENTER to save the consolidated files here:" -ForegroundColor DarkGray
    Write-Host $SourceDirectory -ForegroundColor DarkGray
    Write-Host ""


    $OutputDirectory = Read-Host "Enter a different output directory or press ENTER"

    $OutputDirectory = $OutputDirectory.Trim().Trim('"')


    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {

        $OutputDirectory = $SourceDirectory

    }
    else {

        if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {

            Write-Host ""
            Write-Host "Output directory does not exist." -ForegroundColor Yellow

            $CreateFolder = Read-Host "Create it? (Y/N)"


            if ($CreateFolder -match '^[Yy]') {

                New-Item `
                    -ItemType Directory `
                    -Path $OutputDirectory `
                    -Force | Out-Null

            }
            else {

                Write-Host ""
                Write-Host "Using the source directory instead." -ForegroundColor Yellow

                $OutputDirectory = $SourceDirectory
            }
        }


        if (Test-Path -LiteralPath $OutputDirectory -PathType Container) {

            $OutputDirectory = (
                Resolve-Path -LiteralPath $OutputDirectory
            ).Path
        }
    }


    # ========================================================
    # ALL TRENDLOGS OUTPUT SUBDIRECTORY
    # ========================================================

    $TrendlogOutputDirectory = Join-Path `
        $OutputDirectory `
        "###ALL TRENDLOGS"


    if (-not (Test-Path -LiteralPath $TrendlogOutputDirectory -PathType Container)) {

        New-Item `
            -ItemType Directory `
            -Path $TrendlogOutputDirectory `
            -Force | Out-Null
    }


    $TrendlogOutputDirectory = (
        Resolve-Path -LiteralPath $TrendlogOutputDirectory
    ).Path


    # ========================================================
    # FIND ALL QUALIFYING CSV FILES
    # ========================================================

    Write-Host ""
    Write-Host "Searching recursively for trendlog CSV files..." -ForegroundColor Cyan
    Write-Host ""


    $CsvFiles = @(
        Get-ChildItem `
            -LiteralPath $SourceDirectory `
            -Filter "*.csv" `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match '^Dev\s+(\d+),'
        }
    )


    if ($CsvFiles.Count -eq 0) {

        Write-Host "No qualifying trendlog CSV files were found." -ForegroundColor Red
        Write-Host ""
        Write-Host "Expected filename example:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Dev 10423, BI 6, present-value, UV-3, OCCUPANCY_SNSR-M-2026-09.csv"
        Write-Host ""
    }
    else {

        Write-Host "Trendlog CSV files found: $($CsvFiles.Count)" -ForegroundColor Green
        Write-Host ""


        # ====================================================
        # DEVICE SELECTION
        # ====================================================

        Write-Host "Device selection:" -ForegroundColor Yellow
        Write-Host "Press ENTER to process ALL device instances." -ForegroundColor DarkGray
        Write-Host ""


        $SelectedDevice = Read-Host "Enter a device instance to process only that device, or press ENTER for ALL"

        $SelectedDevice = $SelectedDevice.Trim()


        # ====================================================
        # FILTER TO ONE DEVICE IF REQUESTED
        # ====================================================

        if (-not [string]::IsNullOrWhiteSpace($SelectedDevice)) {

            if ($SelectedDevice -notmatch '^\d+$') {

                Write-Host ""
                Write-Host "Invalid device instance." -ForegroundColor Red
                Write-Host "Device instance must contain numbers only." -ForegroundColor Yellow
                Write-Host ""

                $CsvFiles = @()

            }
            else {

                $CsvFiles = @(
                    $CsvFiles |
                    Where-Object {

                        if ($_.Name -match '^Dev\s+(\d+),') {

                            $Matches[1] -eq $SelectedDevice
                        }
                    }
                )


                if ($CsvFiles.Count -eq 0) {

                    Write-Host ""
                    Write-Host "No trendlog CSV files were found for Device $SelectedDevice." -ForegroundColor Red
                    Write-Host ""

                }
                else {

                    Write-Host ""
                    Write-Host "Processing Device $SelectedDevice only." -ForegroundColor Cyan
                    Write-Host "Trendlog CSV files found: $($CsvFiles.Count)" -ForegroundColor Green
                    Write-Host ""
                }
            }

        }
        else {

            Write-Host ""
            Write-Host "Processing ALL device instances." -ForegroundColor Cyan
            Write-Host ""
        }


        # ====================================================
        # CONTINUE ONLY IF FILES REMAIN
        # ====================================================

        if ($CsvFiles.Count -gt 0) {


            # ====================================================
            # GROUP FILES BY DEVICE INSTANCE
            # ====================================================

            $DeviceGroups = @(
                $CsvFiles |
                ForEach-Object {

                    if ($_.Name -match '^Dev\s+(\d+),') {

                        [PSCustomObject]@{
                            DeviceInstance = $Matches[1]
                            File           = $_
                        }
                    }

                } |
                Group-Object DeviceInstance |
                Sort-Object {
                    [int64]$_.Name
                }
            )


            Write-Host "Unique device instances to process: $($DeviceGroups.Count)" -ForegroundColor Green
            Write-Host ""


            # ====================================================
            # START EXCEL
            # ====================================================

            $Excel = $null

            Write-Host "Starting Microsoft Excel..." -ForegroundColor Cyan
            Write-Host ""


            try {

                $Excel = New-Object -ComObject Excel.Application

                $Excel.Visible = $false
                $Excel.DisplayAlerts = $false
                $Excel.ScreenUpdating = $false


                # XLSX format
                $xlOpenXMLWorkbook = 51


                $OverallStart = Get-Date

                $DeviceNumber = 0
                $DevicesCreated = 0
                $DevicesSkipped = 0

                # Track ONLY workbooks created during this run.
                # These paths are used by the optional master consolidation step.
                $CreatedWorkbookPaths = New-Object System.Collections.Generic.List[string]


                # ====================================================
                # PROCESS EACH DEVICE
                # ====================================================

                foreach ($DeviceGroup in $DeviceGroups) {

                    $DeviceNumber++

                    $DeviceInstance = $DeviceGroup.Name


                    $Files = @(
                        $DeviceGroup.Group.File |
                        Sort-Object Name
                    )


                    Write-Host "============================================================" -ForegroundColor DarkCyan
                    Write-Host "[$DeviceNumber/$($DeviceGroups.Count)] Device $DeviceInstance" -ForegroundColor Cyan
                    Write-Host "Trendlogs: $($Files.Count)" -ForegroundColor White
                    Write-Host "============================================================" -ForegroundColor DarkCyan


                    # ====================================================
                    # MASTER DATA COLLECTION
                    # ====================================================

                    $RowsByTime = @{}

                    $TrendNames = New-Object System.Collections.Generic.List[string]

                    # Unit label used as the worksheet name (for example RTU-1, VAV-2.3, FCU-2).
                    # It is taken from the filename segment after "present-value".
                    $DeviceUnitLabel = $null

                    $TrendCounter = 0


                    # ====================================================
                    # READ EACH CSV FOR THIS DEVICE
                    # ====================================================

                    foreach ($CsvFile in $Files) {

                        $TrendCounter++


                        Write-Host ""
                        Write-Host "  [$TrendCounter/$($Files.Count)] $($CsvFile.Name)" -ForegroundColor Gray


                        # ================================================
                        # IMPORT CSV
                        # ================================================

                        try {

                            $CsvData = @(
                                Import-Csv -LiteralPath $CsvFile.FullName
                            )

                        }
                        catch {

                            Write-Host "      ERROR reading CSV. Skipping." -ForegroundColor Red
                            Write-Host "      $($_.Exception.Message)" -ForegroundColor DarkRed

                            continue
                        }


                        if ($CsvData.Count -eq 0) {

                            Write-Host "      CSV contains no data. Skipping." -ForegroundColor Yellow

                            continue
                        }


                        # ================================================
                        # IDENTIFY COLUMNS
                        # ================================================

                        $Headers = @(
                            $CsvData[0].PSObject.Properties.Name
                        )


                        # Find Time column

                        $TimeColumn = $Headers |
                            Where-Object {
                                $_ -match '^(Time|DateTime|Timestamp)$'
                            } |
                            Select-Object -First 1


                        if (-not $TimeColumn) {

                            $TimeColumn = $Headers |
                                Where-Object {
                                    $_ -match 'Time'
                                } |
                                Select-Object -First 1
                        }


                        if (-not $TimeColumn) {

                            Write-Host "      No Time column found. Skipping." -ForegroundColor Yellow

                            continue
                        }


                        # Find trend value column.
                        # Ignore Time and Events.

                        $ValueColumns = @(
                            $Headers |
                            Where-Object {

                                $_ -ne $TimeColumn -and
                                $_ -notmatch '^Events?$'

                            }
                        )


                        if ($ValueColumns.Count -eq 0) {

                            Write-Host "      No trend value column found. Skipping." -ForegroundColor Yellow

                            continue
                        }


                        $ValueColumn = $ValueColumns[0]


                        # ================================================
                        # POINT / UNIT / DESCRIPTION FROM FILE + CSV HEADER
                        # ================================================

                        # Expected filename pattern:
                        # Dev 10423, BI 6, present-value, UV-3, DESCRIPTION-M-2026-09.csv
                        #
                        # The BACnet point name ALWAYS follows the device instance.
                        # Examples:
                        #   BI 6  -> BI-6
                        #   AV 2  -> AV-2
                        #   BV 26 -> BV-26
                        #   MV 30 -> MV-30

                        $PointName = $null
                        $UnitLabelFromFile = $null


                        if (
                            $CsvFile.BaseName -match
                            '^Dev\s+\d+\s*,\s*([^,]+)\s*,\s*[^,]+\s*,\s*([^,]+)(?:,\s*(.*))?$'
                        ) {

                            $PointNameRaw = $Matches[1].Trim()
                            $UnitLabelFromFile = $Matches[2].Trim()


                            # Normalize "AV 2" / "AV-2" to "AV-2".
                            if ($PointNameRaw -match '^([A-Za-z]+)[\s-]+(\d+)$') {

                                $PointName = "$($Matches[1].ToUpperInvariant())-$($Matches[2])"

                            }
                            else {

                                $PointName = $PointNameRaw
                            }


                            # Normalize common unit labels such as RTU1 / RTU 1 / RTU-1
                            # to RTU-1. Labels already containing more complex numbering
                            # such as HP-1-29 are left unchanged.
                            if (-not [string]::IsNullOrWhiteSpace($UnitLabelFromFile)) {

                                if ($UnitLabelFromFile -match '^([A-Za-z]+)[\s-]*(\d+(?:\.\d+)*)$') {

                                    $NormalizedUnitLabel = "$($Matches[1].ToUpperInvariant())-$($Matches[2])"

                                }
                                else {

                                    $NormalizedUnitLabel = $UnitLabelFromFile
                                }


                                if ([string]::IsNullOrWhiteSpace($DeviceUnitLabel)) {

                                    $DeviceUnitLabel = $NormalizedUnitLabel
                                }
                            }
                        }


                        # Fallback if a filename does not match the expected pattern.
                        if ([string]::IsNullOrWhiteSpace($PointName)) {

                            $PointName = $CsvFile.BaseName
                        }


                        # The original CSV value header commonly looks like:
                        #   RTU1, HTG OUTPUT CMD
                        #
                        # Keep only the description portion after the first comma.
                        # If there is no description, the column header is point name only.
                        $PointDescription = $ValueColumn.Trim()


                        if ($PointDescription -match '^[^,]+,\s*(.+)$') {

                            $PointDescription = $Matches[1].Trim()
                        }


                        if (
                            [string]::IsNullOrWhiteSpace($PointDescription) -or
                            $PointDescription -eq $PointName -or
                            $PointDescription -eq $UnitLabelFromFile -or
                            $PointDescription -eq $DeviceUnitLabel
                        ) {

                            $TrendName = $PointName

                        }
                        else {

                            $TrendName = "$PointName, $PointDescription"
                        }


                        # Make duplicate trend names unique

                        $OriginalTrendName = $TrendName

                        $DuplicateNumber = 2


                        while ($TrendNames -contains $TrendName) {

                            $TrendName = "$OriginalTrendName ($DuplicateNumber)"

                            $DuplicateNumber++
                        }


                        $TrendNames.Add($TrendName)


                        Write-Host "      Trend: $TrendName" -ForegroundColor DarkGray
                        Write-Host "      Rows : $($CsvData.Count)" -ForegroundColor DarkGray


                        # ================================================
                        # READ DATA ROWS
                        # ================================================

                        foreach ($Row in $CsvData) {

                            $TimeText = [string]$Row.$TimeColumn


                            if ([string]::IsNullOrWhiteSpace($TimeText)) {
                                continue
                            }


                            # --------------------------------------------
                            # Parse timestamp
                            # --------------------------------------------

                            $ParsedDate = [datetime]::MinValue


                            $DateParsed = [datetime]::TryParse(
                                $TimeText,
                                [ref]$ParsedDate
                            )


                            if ($DateParsed) {

                                # Identical timestamps from separate
                                # trendlogs will share the same Excel row.

                                $TimeKey = "DATE_$($ParsedDate.Ticks)"

                            }
                            else {

                                # Fallback for unexpected time formats.

                                $TimeKey = "TEXT_$TimeText"
                            }


                            # --------------------------------------------
                            # Create timestamp row if needed
                            # --------------------------------------------

                            if (-not $RowsByTime.ContainsKey($TimeKey)) {

                                $RowsByTime[$TimeKey] = [PSCustomObject]@{

                                    ParsedTime = if ($DateParsed) {
                                        $ParsedDate
                                    }
                                    else {
                                        $null
                                    }

                                    TimeText = $TimeText

                                    Values = @{}
                                }
                            }


                            # --------------------------------------------
                            # Store trend value
                            # --------------------------------------------

                            $TrendValue = $Row.$ValueColumn


                            $RowsByTime[$TimeKey].Values[$TrendName] = $TrendValue
                        }
                    }


                    # ====================================================
                    # VERIFY USABLE DATA
                    # ====================================================

                    if ($TrendNames.Count -eq 0) {

                        Write-Host ""
                        Write-Host "No usable trend data found for Device $DeviceInstance." -ForegroundColor Yellow
                        Write-Host ""

                        $DevicesSkipped++

                        continue
                    }


                    # ====================================================
                    # SORT TIMESTAMPS
                    # ====================================================

                    $SortedRows = @(
                        $RowsByTime.Values |
                        Sort-Object `
                            @{
                                Expression = {

                                    if ($null -ne $_.ParsedTime) {

                                        $_.ParsedTime

                                    }
                                    else {

                                        [datetime]::MaxValue
                                    }
                                }
                            },
                            @{
                                Expression = {
                                    $_.TimeText
                                }
                            }
                    )


                    Write-Host ""
                    Write-Host "  Combined timestamps: $($SortedRows.Count)" -ForegroundColor Green
                    Write-Host "  Combined trends    : $($TrendNames.Count)" -ForegroundColor Green
                    Write-Host ""


                    # ====================================================
                    # CREATE EXCEL WORKBOOK
                    # ====================================================

                    $Workbook = $null
                    $Worksheet = $null
                    $DataRange = $null


                    try {

                        $Workbook = $Excel.Workbooks.Add()


                        # ================================================
                        # KEEP ONLY ONE WORKSHEET
                        # ================================================

                        $Worksheet = $Workbook.Worksheets.Item(1)


                        # Use the unit label as the worksheet name.
                        # Fall back to the device instance if a unit label was not found.
                        $WorksheetName = $DeviceUnitLabel


                        if ([string]::IsNullOrWhiteSpace($WorksheetName)) {

                            $WorksheetName = "Dev $DeviceInstance"
                        }


                        # Excel worksheet names cannot contain these characters
                        # and are limited to 31 characters.
                        $WorksheetName = $WorksheetName -replace '[\\/:?*\[\]]', '_'
                        $WorksheetName = $WorksheetName.Trim()


                        if ($WorksheetName.Length -gt 31) {

                            $WorksheetName = $WorksheetName.Substring(0, 31)
                        }


                        $Worksheet.Name = $WorksheetName


                        # Delete any additional default worksheets

                        while ($Workbook.Worksheets.Count -gt 1) {

                            $Workbook.Worksheets.Item(
                                $Workbook.Worksheets.Count
                            ).Delete()
                        }


                        # ================================================
                        # DIMENSIONS
                        # ================================================

                        $TotalColumns = $TrendNames.Count + 1

                        $TotalDataRows = $SortedRows.Count

                        $TotalExcelRows = $TotalDataRows + 1


                        # ================================================
                        # WRITE HEADERS
                        # ================================================

                        $Worksheet.Cells.Item(1, 1).Value2 = "Time"


                        for (
                            $ColumnIndex = 0;
                            $ColumnIndex -lt $TrendNames.Count;
                            $ColumnIndex++
                        ) {

                            $ExcelColumn = $ColumnIndex + 2


                            $Worksheet.Cells.Item(
                                1,
                                $ExcelColumn
                            ).Value2 = [string]$TrendNames[$ColumnIndex]
                        }


                        # ================================================
                        # BUILD 2D DATA ARRAY
                        # ================================================

                        if ($TotalDataRows -gt 0) {

                            $DataArray = New-Object 'object[,]' `
                                $TotalDataRows, $TotalColumns


                            for (
                                $RowIndex = 0;
                                $RowIndex -lt $TotalDataRows;
                                $RowIndex++
                            ) {

                                $CurrentRow = $SortedRows[$RowIndex]


                                # ----------------------------------------
                                # TIME
                                # ----------------------------------------

                                if ($null -ne $CurrentRow.ParsedTime) {

                                    $ExcelDate = $CurrentRow.ParsedTime.ToOADate()


                                    $DataArray.SetValue(
                                        $ExcelDate,
                                        $RowIndex,
                                        0
                                    )

                                }
                                else {

                                    $DataArray.SetValue(
                                        [string]$CurrentRow.TimeText,
                                        $RowIndex,
                                        0
                                    )
                                }


                                # ----------------------------------------
                                # TREND VALUES
                                # ----------------------------------------

                                for (
                                    $TrendIndex = 0;
                                    $TrendIndex -lt $TrendNames.Count;
                                    $TrendIndex++
                                ) {

                                    $CurrentTrend = $TrendNames[$TrendIndex]

                                    $ArrayColumn = $TrendIndex + 1


                                    if (
                                        $CurrentRow.Values.ContainsKey(
                                            $CurrentTrend
                                        )
                                    ) {

                                        $Value = $CurrentRow.Values[$CurrentTrend]


                                        # --------------------------------
                                        # Convert numeric values
                                        # --------------------------------

                                        $NumericValue = 0.0


                                        $IsNumeric = [double]::TryParse(

                                            [string]$Value,

                                            [System.Globalization.NumberStyles]::Any,

                                            [System.Globalization.CultureInfo]::InvariantCulture,

                                            [ref]$NumericValue
                                        )


                                        if ($IsNumeric) {

                                            $DataArray.SetValue(
                                                $NumericValue,
                                                $RowIndex,
                                                $ArrayColumn
                                            )

                                        }
                                        else {

                                            $DataArray.SetValue(
                                                [string]$Value,
                                                $RowIndex,
                                                $ArrayColumn
                                            )
                                        }

                                    }
                                    else {

                                        $DataArray.SetValue(
                                            $null,
                                            $RowIndex,
                                            $ArrayColumn
                                        )
                                    }
                                }
                            }


                            # ============================================
                            # WRITE DATA TO EXCEL
                            # ============================================

                            $StartCell = $Worksheet.Cells.Item(2, 1)


                            $EndCell = $Worksheet.Cells.Item(
                                $TotalExcelRows,
                                $TotalColumns
                            )


                            $DataRange = $Worksheet.Range(
                                $StartCell,
                                $EndCell
                            )


                            $DataRange.Value2 = $DataArray
                        }


                        # ================================================
                        # FORMAT HEADER
                        # ================================================

                        $HeaderRange = $Worksheet.Range(
                            $Worksheet.Cells.Item(1, 1),
                            $Worksheet.Cells.Item(
                                1,
                                $TotalColumns
                            )
                        )


                        $HeaderRange.Font.Bold = $true


                        # ================================================
                        # FORMAT TIME COLUMN
                        # ================================================

                        if ($TotalDataRows -gt 0) {

                            $TimeRange = $Worksheet.Range(
                                $Worksheet.Cells.Item(2, 1),
                                $Worksheet.Cells.Item(
                                    $TotalExcelRows,
                                    1
                                )
                            )


                            $TimeRange.NumberFormat = "m/d/yyyy h:mm"
                        }


                        # ================================================
                        # AUTOFILTER
                        # ================================================

                        $FilterRange = $Worksheet.Range(
                            $Worksheet.Cells.Item(1, 1),
                            $Worksheet.Cells.Item(
                                $TotalExcelRows,
                                $TotalColumns
                            )
                        )


                        $FilterRange.AutoFilter() | Out-Null


                        # ================================================
                        # FREEZE TOP ROW
                        # ================================================

                        $Worksheet.Activate()

                        $Excel.ActiveWindow.SplitRow = 1

                        $Excel.ActiveWindow.FreezePanes = $true


                        # ================================================
                        # COLUMN WIDTHS
                        # ================================================

                        $Worksheet.UsedRange.Columns.AutoFit() | Out-Null


                        for (
                            $ColumnNumber = 1;
                            $ColumnNumber -le $TotalColumns;
                            $ColumnNumber++
                        ) {

                            $CurrentColumn = $Worksheet.Columns.Item(
                                $ColumnNumber
                            )


                            if ($CurrentColumn.ColumnWidth -gt 40) {

                                $CurrentColumn.ColumnWidth = 40
                            }
                        }


                        # Make timestamp column readable

                        if ($Worksheet.Columns.Item(1).ColumnWidth -lt 18) {

                            $Worksheet.Columns.Item(1).ColumnWidth = 18
                        }


                        # ====================================================
                        # OUTPUT FILE NAME
                        # ====================================================

                        $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"


                        # Include the unit label directly after the device instance
                        # in the consolidated workbook filename.
                        # Example: Dev 10423 RTU-1 ALL TRENDLOGS 20260914_112500.xlsx
                        $FileUnitLabel = $DeviceUnitLabel


                        if ([string]::IsNullOrWhiteSpace($FileUnitLabel)) {

                            $FileUnitLabel = "UNKNOWN UNIT"
                        }


                        # Remove characters that are invalid in Windows filenames.
                        $FileUnitLabel = $FileUnitLabel -replace '[<>:"/\|?*]', '_'
                        $FileUnitLabel = $FileUnitLabel.Trim()


                        $OutputFileName = (
                            "Dev {0} {1} ALL TRENDLOGS {2}.xlsx" -f `
                            $DeviceInstance,
                            $FileUnitLabel,
                            $Timestamp
                        )


                        $OutputPath = Join-Path `
                            $TrendlogOutputDirectory `
                            $OutputFileName


                        # Additional duplicate protection

                        $DuplicateCounter = 2


                        while (Test-Path -LiteralPath $OutputPath) {

                            $OutputFileName = (
                                "Dev {0} {1} ALL TRENDLOGS {2} ({3}).xlsx" -f `
                                $DeviceInstance,
                                $FileUnitLabel,
                                $Timestamp,
                                $DuplicateCounter
                            )


                            $OutputPath = Join-Path `
                                $TrendlogOutputDirectory `
                                $OutputFileName


                            $DuplicateCounter++
                        }


                        # ====================================================
                        # SAVE WORKBOOK
                        # ====================================================

                        $Workbook.SaveAs(
                            $OutputPath,
                            $xlOpenXMLWorkbook
                        )


                        $DevicesCreated++
                        $CreatedWorkbookPaths.Add($OutputPath)


                        Write-Host "  SAVED:" -ForegroundColor Green
                        Write-Host "  $OutputPath" -ForegroundColor White
                        Write-Host ""

                    }
                    catch {

                        $DevicesSkipped++

                        Write-Host ""
                        Write-Host "ERROR processing Device $DeviceInstance" -ForegroundColor Red
                        Write-Host $_.Exception.Message -ForegroundColor Red
                        Write-Host ""

                    }
                    finally {

                        # ====================================================
                        # CLOSE WORKBOOK
                        # ====================================================

                        if ($Workbook) {

                            try {

                                $Workbook.Close($false)

                            }
                            catch {
                            }


                            try {

                                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                                    $Workbook
                                )

                            }
                            catch {
                            }


                            $Workbook = $null
                        }


                        if ($Worksheet) {

                            try {

                                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                                    $Worksheet
                                )

                            }
                            catch {
                            }


                            $Worksheet = $null
                        }


                        if ($DataRange) {

                            try {

                                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                                    $DataRange
                                )

                            }
                            catch {
                            }


                            $DataRange = $null
                        }
                    }
                }


                # ====================================================
                # OPTIONAL MASTER WORKBOOK CONSOLIDATION
                # ====================================================

                if ($CreatedWorkbookPaths.Count -gt 0) {

                    Write-Host ""
                    Write-Host "============================================================" -ForegroundColor Cyan
                    Write-Host " OPTIONAL MASTER WORKBOOK" -ForegroundColor Cyan
                    Write-Host "============================================================" -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "The workbooks created during this run can now be combined" -ForegroundColor Yellow
                    Write-Host "into one workbook. Each worksheet will remain on its own tab." -ForegroundColor Yellow
                    Write-Host "Unit labels will be retained as the consolidated tab names." -ForegroundColor Yellow
                    Write-Host ""

                    $CreateMasterWorkbook = Read-Host "Consolidate all newly created workbooks into one workbook? (Y/n) [Default: Y]"

                    # Default is YES: pressing Enter consolidates. Only an explicit N skips it.
                    if ($CreateMasterWorkbook -notmatch '^[Nn]') {

                        $MasterWorkbook = $null
                        $MasterWorksheet = $null
                        $SourceWorkbook = $null
                        $SourceWorksheet = $null
                        $MasterOutputPath = $null

                        try {

                            Write-Host ""
                            Write-Host "Creating consolidated master workbook..." -ForegroundColor Cyan
                            Write-Host ""

                            $MasterWorkbook = $Excel.Workbooks.Add()

                            # Keep one temporary sheet until source sheets have been copied.
                            $MasterWorksheet = $MasterWorkbook.Worksheets.Item(1)
                            $MasterWorksheet.Name = "TEMP"

                            while ($MasterWorkbook.Worksheets.Count -gt 1) {
                                $MasterWorkbook.Worksheets.Item($MasterWorkbook.Worksheets.Count).Delete()
                            }

                            $CopiedSheetCount = 0
                            $SourceNumber = 0

                            foreach ($CreatedWorkbookPath in $CreatedWorkbookPaths) {

                                $SourceNumber++

                                if (-not (Test-Path -LiteralPath $CreatedWorkbookPath -PathType Leaf)) {
                                    Write-Host "  [$SourceNumber/$($CreatedWorkbookPaths.Count)] Missing file. Skipping:" -ForegroundColor Yellow
                                    Write-Host "      $CreatedWorkbookPath" -ForegroundColor DarkYellow
                                    continue
                                }

                                $Extension = [System.IO.Path]::GetExtension($CreatedWorkbookPath).ToLowerInvariant()

                                if ($Extension -notin @('.csv', '.xls', '.xlsx', '.xlsm', '.xlsb')) {
                                    Write-Host "  [$SourceNumber/$($CreatedWorkbookPaths.Count)] Unsupported spreadsheet type. Skipping:" -ForegroundColor Yellow
                                    Write-Host "      $CreatedWorkbookPath" -ForegroundColor DarkYellow
                                    continue
                                }

                                Write-Host "  [$SourceNumber/$($CreatedWorkbookPaths.Count)] $([System.IO.Path]::GetFileName($CreatedWorkbookPath))" -ForegroundColor Gray

                                try {

                                    $SourceWorkbook = $Excel.Workbooks.Open($CreatedWorkbookPath)

                                    # Device instance comes from the source filename when available.
                                    $DeviceInstanceForTabs = $null
                                    $SourceBaseName = [System.IO.Path]::GetFileNameWithoutExtension($CreatedWorkbookPath)

                                    if ($SourceBaseName -match '(?i)\bDev\s+(\d+)') {
                                        $DeviceInstanceForTabs = $Matches[1]
                                    }

                                    for ($SheetIndex = 1; $SheetIndex -le $SourceWorkbook.Worksheets.Count; $SheetIndex++) {

                                        $SourceWorksheet = $SourceWorkbook.Worksheets.Item($SheetIndex)
                                        $OriginalSheetName = [string]$SourceWorksheet.Name

                                        # Copy the full worksheet so formatting, filters, widths,
                                        # freeze panes, formulas, and data remain intact.
                                        $SourceWorksheet.Copy(
                                            [Type]::Missing,
                                            $MasterWorkbook.Worksheets.Item($MasterWorkbook.Worksheets.Count)
                                        )

                                        $CopiedSheet = $MasterWorkbook.Worksheets.Item($MasterWorkbook.Worksheets.Count)

                                        try {

                                            # Preserve the source worksheet name (the unit label)
                                            # in the consolidated master workbook.
                                            $DesiredSheetName = $OriginalSheetName

                                            # Excel worksheet names cannot contain these characters.
                                            $DesiredSheetName = $DesiredSheetName -replace '[\\/:?*\[\]]', '_'
                                            $DesiredSheetName = $DesiredSheetName.Trim()

                                            if ([string]::IsNullOrWhiteSpace($DesiredSheetName)) {
                                                $DesiredSheetName = "Sheet"
                                            }

                                            # Excel worksheet names are limited to 31 characters.
                                            if ($DesiredSheetName.Length -gt 31) {
                                                $DesiredSheetName = $DesiredSheetName.Substring(0, 31)
                                            }

                                            $FinalSheetName = $DesiredSheetName
                                            $SheetNameCounter = 2

                                            while (@($MasterWorkbook.Worksheets | ForEach-Object { $_.Name }) -contains $FinalSheetName) {

                                                $Suffix = " ($SheetNameCounter)"
                                                $MaximumBaseLength = 31 - $Suffix.Length

                                                if ($DesiredSheetName.Length -gt $MaximumBaseLength) {
                                                    $FinalSheetName = $DesiredSheetName.Substring(0, $MaximumBaseLength) + $Suffix
                                                }
                                                else {
                                                    $FinalSheetName = $DesiredSheetName + $Suffix
                                                }

                                                $SheetNameCounter++
                                            }

                                            $CopiedSheet.Name = $FinalSheetName
                                            $CopiedSheetCount++

                                            Write-Host "      Tab: $FinalSheetName" -ForegroundColor DarkGray

                                        }
                                        finally {

                                            if ($CopiedSheet) {
                                                try {
                                                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($CopiedSheet)
                                                }
                                                catch {
                                                }
                                                $CopiedSheet = $null
                                            }
                                        }

                                        if ($SourceWorksheet) {
                                            try {
                                                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($SourceWorksheet)
                                            }
                                            catch {
                                            }
                                            $SourceWorksheet = $null
                                        }
                                    }

                                }
                                catch {

                                    Write-Host "      ERROR opening/copying source workbook. Skipping." -ForegroundColor Red
                                    Write-Host "      $($_.Exception.Message)" -ForegroundColor DarkRed

                                }
                                finally {

                                    if ($SourceWorkbook) {
                                        try {
                                            $SourceWorkbook.Close($false)
                                        }
                                        catch {
                                        }

                                        try {
                                            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($SourceWorkbook)
                                        }
                                        catch {
                                        }

                                        $SourceWorkbook = $null
                                    }

                                    if ($SourceWorksheet) {
                                        try {
                                            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($SourceWorksheet)
                                        }
                                        catch {
                                        }

                                        $SourceWorksheet = $null
                                    }
                                }
                            }

                            if ($CopiedSheetCount -gt 0) {

                                # Remove the temporary worksheet after at least one real sheet exists.
                                $MasterWorksheet.Delete()

                                try {
                                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($MasterWorksheet)
                                }
                                catch {
                                }
                                $MasterWorksheet = $null

                                $MasterTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                                $MasterOutputFileName = "ALL DEVICES CONSOLIDATED TRENDLOGS $MasterTimestamp.xlsx"
                                $MasterOutputPath = Join-Path $TrendlogOutputDirectory $MasterOutputFileName

                                $MasterDuplicateCounter = 2

                                while (Test-Path -LiteralPath $MasterOutputPath) {

                                    $MasterOutputFileName = "ALL DEVICES CONSOLIDATED TRENDLOGS $MasterTimestamp ($MasterDuplicateCounter).xlsx"
                                    $MasterOutputPath = Join-Path $TrendlogOutputDirectory $MasterOutputFileName
                                    $MasterDuplicateCounter++
                                }

                                $MasterWorkbook.SaveAs(
                                    $MasterOutputPath,
                                    $xlOpenXMLWorkbook
                                )

                                Write-Host ""
                                Write-Host "MASTER WORKBOOK SAVED:" -ForegroundColor Green
                                Write-Host $MasterOutputPath -ForegroundColor White
                                Write-Host "Worksheets copied: $CopiedSheetCount" -ForegroundColor Green
                                Write-Host ""

                            }
                            else {

                                Write-Host ""
                                Write-Host "No worksheets were successfully copied. Master workbook was not saved." -ForegroundColor Yellow
                                Write-Host ""
                            }

                        }
                        catch {

                            Write-Host ""
                            Write-Host "ERROR creating master workbook:" -ForegroundColor Red
                            Write-Host $_.Exception.Message -ForegroundColor Red
                            Write-Host ""

                        }
                        finally {

                            if ($SourceWorkbook) {
                                try {
                                    $SourceWorkbook.Close($false)
                                }
                                catch {
                                }

                                try {
                                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($SourceWorkbook)
                                }
                                catch {
                                }

                                $SourceWorkbook = $null
                            }

                            if ($SourceWorksheet) {
                                try {
                                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($SourceWorksheet)
                                }
                                catch {
                                }

                                $SourceWorksheet = $null
                            }

                            if ($MasterWorksheet) {
                                try {
                                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($MasterWorksheet)
                                }
                                catch {
                                }

                                $MasterWorksheet = $null
                            }

                            if ($MasterWorkbook) {
                                try {
                                    $MasterWorkbook.Close($false)
                                }
                                catch {
                                }

                                try {
                                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($MasterWorkbook)
                                }
                                catch {
                                }

                                $MasterWorkbook = $null
                            }
                        }
                    }
                    else {

                        Write-Host ""
                        Write-Host "Master workbook consolidation skipped." -ForegroundColor DarkGray
                        Write-Host ""
                    }
                }


                # ====================================================
                # COMPLETION SUMMARY
                # ====================================================

                $Elapsed = (Get-Date) - $OverallStart


                Write-Host ""
                Write-Host "============================================================" -ForegroundColor Green
                Write-Host " COMPLETE" -ForegroundColor Green
                Write-Host "============================================================" -ForegroundColor Green
                Write-Host ""


                if (-not [string]::IsNullOrWhiteSpace($SelectedDevice)) {

                    Write-Host "Device selection        : $SelectedDevice"

                }
                else {

                    Write-Host "Device selection        : ALL"
                }


                Write-Host "Device instances found : $($DeviceGroups.Count)"
                Write-Host "Workbooks created      : $DevicesCreated"


                if ($DevicesSkipped -gt 0) {

                    Write-Host "Devices skipped/errors : $DevicesSkipped" -ForegroundColor Yellow

                }
                else {

                    Write-Host "Devices skipped/errors : 0"
                }


                Write-Host ""
                Write-Host "Output directory:" -ForegroundColor Yellow
                Write-Host $TrendlogOutputDirectory
                Write-Host ""


                Write-Host (
                    "Elapsed time: {0:hh\:mm\:ss}" -f $Elapsed
                )


                Write-Host ""

            }
            catch {

                Write-Host ""
                Write-Host "A fatal error occurred:" -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor Red
                Write-Host ""

            }
            finally {

                # ====================================================
                # CLOSE EXCEL COMPLETELY
                # ====================================================

                if ($Excel) {

                    try {

                        $Excel.Quit()

                    }
                    catch {
                    }


                    try {

                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                            $Excel
                        )

                    }
                    catch {
                    }


                    $Excel = $null
                }


                # COM cleanup

                [GC]::Collect()

                [GC]::WaitForPendingFinalizers()

                [GC]::Collect()

                [GC]::WaitForPendingFinalizers()
            }
        }
    }


    # ========================================================
    # RUN AGAIN
    # ========================================================

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""

    $RunAgain = Read-Host "Enter anything to run again, or press ENTER to exit"

}
while (-not [string]::IsNullOrWhiteSpace($RunAgain))


Write-Host ""
Write-Host "Exiting." -ForegroundColor Cyan
Write-Host ""
