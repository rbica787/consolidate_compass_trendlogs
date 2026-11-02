#Requires -Version 5.1
# ============================================================
# CONSOLIDATE TRENDLOG CSV FILES BY BACNET DEVICE INSTANCE
# ============================================================
# - Searches a selected directory recursively for Alerton CSVs.
# - Skips empty CSV files; optionally processes one device only.
# - Creates one XLSX per device in "###ALL TRENDLOGS".
# - Combines all BACnet points on a single timestamp axis.
# - Merges repeated exports of the SAME point into ONE column.
# - For conflicting readings, the source with more populated,
#   UNIQUE timestamps takes priority (not necessarily newer).
# - Adds distinct timestamps from smaller exports and fills
#   blanks, but never replaces an existing value with a blank.
# - Optionally combines this run's device workbooks into one.
# - Does not change the original CSV files.
# - Uses caching, bulk Excel writes, and fixed column widths.
# - Uses Array.SetValue(row,column) to prevent PowerShell's
#   multidimensional index expression parsing ambiguity.
# Requires Windows with Microsoft Excel installed.
# ============================================================

$ErrorActionPreference = 'Stop'

function Get-TimeInfo {
    param([string]$Text)

    $clean = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }

    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture

    $valid = [datetime]::TryParse($clean, $invariant, $styles, [ref]$parsed)
    if (-not $valid) {
        $valid = [datetime]::TryParse($clean, [ref]$parsed)
    }

    if ($valid) {
        # The DateTime ticks key collapses equivalent date/time strings.
        return [pscustomobject]@{
            Key        = 'DATE_' + $parsed.Ticks
            ParsedTime = $parsed
            TimeText   = $clean
        }
    }

    # Unrecognized timestamps are preserved as text rather than discarded.
    return [pscustomobject]@{
        Key        = 'TEXT_' + $clean
        ParsedTime = $null
        TimeText   = $clean
    }
}

function Get-UnitLabel {
    param([string]$RawLabel)
    $label = $RawLabel.Trim()
    if ($label -match '^([A-Za-z]+)[\s-]*(\d+(?:\.\d+)*)$') {
        return ('{0}-{1}' -f $Matches[1].ToUpperInvariant(), $Matches[2])
    }
    return $label
}

function Get-SafeSheetName {
    param([string]$Name)
    $result = ($Name -replace '[\\/:?*\[\]]', '_').Trim().Trim("'")
    if ([string]::IsNullOrWhiteSpace($result)) { $result = 'Sheet' }
    if ($result.Length -gt 31) { $result = $result.Substring(0,31) }
    return $result
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) }
        catch { }
    }
}

function Get-UniqueSheetName {
    param($Workbook, [string]$Proposed)
    $base = Get-SafeSheetName $Proposed
    $names = @($Workbook.Worksheets | ForEach-Object { [string]$_.Name })
    $candidate = $base
    $counter = 2
    while ($names -contains $candidate) {
        $suffix = " ($counter)"
        $maxBase = 31 - $suffix.Length
        $candidate = $base.Substring(0, [Math]::Min($base.Length,$maxBase)) + $suffix
        $counter++
    }
    return $candidate
}

do {
    Clear-Host
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ' Trendlog Consolidation by Device Instance' -ForegroundColor Cyan
    Write-Host ' Duplicate timestamps / duplicate point exports merged' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''

    do {
        $SourceDirectory = (Read-Host 'Enter the directory containing the trendlog CSV files').Trim().Trim('"')
        if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
            Write-Host 'Directory does not exist. Please try again.' -ForegroundColor Red
            $SourceDirectory = $null
        }
    } until ($SourceDirectory)
    $SourceDirectory = (Resolve-Path -LiteralPath $SourceDirectory).Path

    Write-Host ''
    Write-Host 'Press ENTER to save the consolidated files here:' -ForegroundColor Yellow
    Write-Host $SourceDirectory -ForegroundColor DarkGray
    $OutputDirectory = (Read-Host 'Enter a different output directory or press ENTER').Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $OutputDirectory = $SourceDirectory
    }
    elseif (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        $create = Read-Host 'Output directory does not exist. Create it? (Y/N)'
        if ($create -match '^[Yy]') {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        }
        else {
            Write-Host 'Using the source directory instead.' -ForegroundColor Yellow
            $OutputDirectory = $SourceDirectory
        }
    }
    $OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
    $TrendlogOutputDirectory = Join-Path $OutputDirectory '###ALL TRENDLOGS'
    if (-not (Test-Path -LiteralPath $TrendlogOutputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $TrendlogOutputDirectory -Force | Out-Null
    }

    Write-Host ''
    Write-Host 'Searching recursively for trendlog CSV files...' -ForegroundColor Cyan
    $allCsv = @(Get-ChildItem -LiteralPath $SourceDirectory -Filter '*.csv' -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^Dev\s+(\d+)\s*,' })
    $empty = @($allCsv | Where-Object { $_.Length -eq 0 })
    $CsvFiles = @($allCsv | Where-Object { $_.Length -gt 0 })
    if ($empty.Count -gt 0) {
        Write-Host "Empty trendlog CSV files skipped: $($empty.Count)" -ForegroundColor Yellow
    }
    Write-Host "Nonempty trendlog CSV files found: $($CsvFiles.Count)" -ForegroundColor Green

    if ($CsvFiles.Count -gt 0) {
        Write-Host ''
        $SelectedDevice = (Read-Host 'Enter a device instance to process only that device, or press ENTER for ALL').Trim()
        if ($SelectedDevice -ne '') {
            if ($SelectedDevice -notmatch '^\d+$') {
                Write-Host 'Device instance must contain numbers only.' -ForegroundColor Red
                $CsvFiles = @()
            }
            else {
                $CsvFiles = @($CsvFiles | Where-Object {
                    $_.Name -match '^Dev\s+(\d+)\s*,' -and $Matches[1] -eq $SelectedDevice
                })
            }
        }
    }

    if ($CsvFiles.Count -eq 0) {
        Write-Host 'No qualifying trendlog CSVs found for this selection.' -ForegroundColor Yellow
    }
    else {
        $DeviceGroups = @($CsvFiles | ForEach-Object {
            if ($_.Name -match '^Dev\s+(\d+)\s*,') {
                [pscustomobject]@{ DeviceInstance = $Matches[1]; File = $_ }
            }
        } | Group-Object DeviceInstance | Sort-Object { [int64]$_.Name })

        $Excel = $null
        $CreatedWorkbookPaths = New-Object 'System.Collections.Generic.List[string]'
        $DevicesCreated = 0
        $DevicesSkipped = 0
        $OverallStart = Get-Date

        try {
            Write-Host 'Starting Microsoft Excel...' -ForegroundColor Cyan
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $false
            $Excel.DisplayAlerts = $false
            $Excel.ScreenUpdating = $false
            $Excel.EnableEvents = $false
            # Avoid setting Excel.Calculation here. Excel can reject manual mode
            # (0x800A03EC), particularly with OLAP-backed workbooks.
            # These generated trendlog workbooks contain no formulas, so
            # changing calculation mode offers little performance benefit.
            $xlOpenXMLWorkbook = 51

            $DeviceNumber = 0
            foreach ($DeviceGroup in $DeviceGroups) {
                $DeviceNumber++
                $DeviceInstance = $DeviceGroup.Name
                $Files = @($DeviceGroup.Group.File | Sort-Object FullName)

                Write-Host ''
                Write-Host '============================================================' -ForegroundColor DarkCyan
                Write-Host "[$DeviceNumber/$($DeviceGroups.Count)] Device $DeviceInstance - $($Files.Count) CSV files" -ForegroundColor Cyan
                Write-Host '============================================================' -ForegroundColor DarkCyan

                # Collect each source separately so source COMPLETENESS can be
                # compared before choosing values at overlapping timestamps.
                $Points = @{}
                $DeviceUnitLabel = $null
                $ImportFailures = 0
                $WithinFileDuplicates = 0
                $FilesImported = 0
                # Cache repeated timestamp parsing across exports for this device.
                $TimeInfoCache = @{}

                foreach ($CsvFile in $Files) {
                    Write-Host "  Reading: $($CsvFile.Name)" -ForegroundColor Gray
                    try {
                        $CsvData = @(Import-Csv -LiteralPath $CsvFile.FullName)
                        if ($CsvData.Count -eq 0) {
                            Write-Host '    No data rows. Skipped.' -ForegroundColor Yellow
                            continue
                        }

                        $Headers = @($CsvData[0].PSObject.Properties.Name)
                        $TimeColumn = $Headers | Where-Object { $_ -match '^(Time|DateTime|Timestamp)$' } | Select-Object -First 1
                        if (-not $TimeColumn) {
                            $TimeColumn = $Headers | Where-Object { $_ -match 'Time' } | Select-Object -First 1
                        }
                        if (-not $TimeColumn) {
                            Write-Host '    No timestamp column. Skipped.' -ForegroundColor Yellow
                            continue
                        }
                        $ValueColumn = $Headers | Where-Object {
                            $_ -ne $TimeColumn -and $_ -notmatch '^Events?$'
                        } | Select-Object -First 1
                        if (-not $ValueColumn) {
                            Write-Host '    No trend value column. Skipped.' -ForegroundColor Yellow
                            continue
                        }

                        $PointName = $null
                        $UnitLabelFromFile = $null
                        if ($CsvFile.BaseName -match '^Dev\s+\d+\s*,\s*([^,]+)\s*,\s*[^,]+\s*,\s*([^,]+)(?:,\s*(.*))?$') {
                            $rawPoint = $Matches[1].Trim()
                            $UnitLabelFromFile = Get-UnitLabel $Matches[2]
                            if ($rawPoint -match '^([A-Za-z]+)[\s-]+(\d+)$') {
                                $PointName = '{0}-{1}' -f $Matches[1].ToUpperInvariant(), $Matches[2]
                            }
                            else { $PointName = $rawPoint }
                        }
                        if (-not $PointName) {
                            # Avoid accidentally merging two unrelated filenames.
                            $PointName = $CsvFile.BaseName
                        }
                        if (-not $DeviceUnitLabel -and $UnitLabelFromFile) {
                            $DeviceUnitLabel = $UnitLabelFromFile
                        }

                        $description = ([string]$ValueColumn).Trim()
                        if ($description -match '^[^,]+,\s*(.+)$') {
                            $description = $Matches[1].Trim()
                        }
                        if ([string]::IsNullOrWhiteSpace($description) -or
                            $description -eq $PointName -or
                            $description -eq $UnitLabelFromFile -or
                            $description -eq $DeviceUnitLabel) {
                            $TrendName = $PointName
                        }
                        else { $TrendName = "$PointName, $description" }

                        # A point is identified by its BACnet point name, NOT the
                        # description or the source filename's date suffix.
                        $PointKey = $PointName.ToUpperInvariant()
                        $SourceValues = @{}
                        $NonemptyUniqueTimestamps = 0

                        foreach ($Row in $CsvData) {
                            $rawTime = ([string]$Row.$TimeColumn).Trim()
                            if (-not $TimeInfoCache.ContainsKey($rawTime)) {
                                $TimeInfoCache[$rawTime] = Get-TimeInfo $rawTime
                            }
                            $info = $TimeInfoCache[$rawTime]
                            if ($null -eq $info) { continue }
                            $value = ([string]$Row.$ValueColumn).Trim()

                            if ($SourceValues.ContainsKey($info.Key)) {
                                $WithinFileDuplicates++
                                $previous = $SourceValues[$info.Key]
                                # For repeated timestamps IN the same CSV,
                                # prefer a populated value; if both populated
                                # and different, retain the later row.
                                if ($value -ne '') { $previous.Value = $value }
                            }
                            else {
                                $SourceValues[$info.Key] = [pscustomobject]@{
                                    ParsedTime = $info.ParsedTime
                                    TimeText   = $info.TimeText
                                    Value      = $value
                                }
                            }
                        }

                        foreach ($record in $SourceValues.Values) {
                            if (-not [string]::IsNullOrWhiteSpace($record.Value)) {
                                $NonemptyUniqueTimestamps++
                            }
                        }
                        if ($SourceValues.Count -eq 0) { continue }

                        $source = [pscustomobject]@{
                            File              = $CsvFile.FullName
                            LastWriteTimeUtc  = $CsvFile.LastWriteTimeUtc
                            TrendName         = $TrendName
                            Samples           = $SourceValues
                            PopulatedCount    = $NonemptyUniqueTimestamps
                            TimestampCount    = $SourceValues.Count
                        }
                        if (-not $Points.ContainsKey($PointKey)) {
                            $Points[$PointKey] = [pscustomobject]@{
                                PointName = $PointName
                                Sources   = (New-Object 'System.Collections.Generic.List[object]')
                            }
                        }
                        $Points[$PointKey].Sources.Add($source)
                        $FilesImported++
                        Write-Host "    $PointName : $NonemptyUniqueTimestamps populated / $($SourceValues.Count) unique timestamps" -ForegroundColor DarkGray
                    }
                    catch {
                        $ImportFailures++
                        Write-Host "    ERROR reading CSV: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    finally { $CsvData = $null }
                }

                if ($Points.Count -eq 0) {
                    Write-Host "No usable trend data found for device $DeviceInstance." -ForegroundColor Yellow
                    $DevicesSkipped++
                    continue
                }

                # One column per point and one row per timestamp across points.
                $RowsByTime = @{}
                $TrendNames = New-Object 'System.Collections.Generic.List[string]'
                $Conflicts = 0
                $BlankValuesFilled = 0
                $DuplicatePointExports = 0
                $OverlappingSamples = 0

                foreach ($pointKey in @($Points.Keys | Sort-Object)) {
                    $point = $Points[$pointKey]
                    # Highest count of POPULATED UNIQUE timestamps wins.
                    # Secondary: total timestamps; tertiary: latest file;
                    # final tie: deterministic file path.
                    $sources = @($point.Sources | Sort-Object `
                        @{ Expression = 'PopulatedCount'; Descending = $true }, `
                        @{ Expression = 'TimestampCount'; Descending = $true }, `
                        @{ Expression = 'LastWriteTimeUtc'; Descending = $true }, `
                        @{ Expression = 'File'; Descending = $false })

                    $DuplicatePointExports += [Math]::Max(0, $sources.Count - 1)
                    $TrendName = [string]$sources[0].TrendName
                    $original = $TrendName
                    $suffixNumber = 2
                    while ($TrendNames -contains $TrendName) {
                        $TrendName = "$original ($suffixNumber)"
                        $suffixNumber++
                    }
                    $TrendNames.Add($TrendName)

                    # Process most-complete source FIRST; less-complete files
                    # can add missing timestamps or fill empty values only.
                    $SeenForThisPoint = @{}
                    foreach ($source in $sources) {
                        foreach ($timeKey in $source.Samples.Keys) {
                            $sample = $source.Samples[$timeKey]
                            $WasSeen = $SeenForThisPoint.ContainsKey($timeKey)
                            if ($WasSeen) { $OverlappingSamples++ }
                            else { $SeenForThisPoint[$timeKey] = $true }
                            if (-not $RowsByTime.ContainsKey($timeKey)) {
                                $RowsByTime[$timeKey] = [pscustomobject]@{
                                    ParsedTime = $sample.ParsedTime
                                    TimeText   = $sample.TimeText
                                    Values     = @{}
                                }
                            }
                            $combined = $RowsByTime[$timeKey].Values
                            $newValue = [string]$sample.Value
                            if (-not $combined.ContainsKey($TrendName)) {
                                if (-not [string]::IsNullOrWhiteSpace($newValue)) {
                                    $combined[$TrendName] = $newValue
                                    if ($WasSeen) { $BlankValuesFilled++ }
                                }
                            }
                            else {
                                $oldValue = [string]$combined[$TrendName]
                                if ([string]::IsNullOrWhiteSpace($oldValue) -and
                                    -not [string]::IsNullOrWhiteSpace($newValue)) {
                                    $combined[$TrendName] = $newValue
                                    $BlankValuesFilled++
                                }
                                elseif (-not [string]::IsNullOrWhiteSpace($newValue) -and
                                    $oldValue -ne $newValue) {
                                    # Keep the value already selected from the
                                    # more complete (higher-priority) source.
                                    $Conflicts++
                                }
                            }
                        }
                    }
                }

                $SortedRows = @($RowsByTime.Values | Sort-Object `
                    @{ Expression = { if ($null -ne $_.ParsedTime) { $_.ParsedTime } else { [datetime]::MaxValue } } }, `
                    @{ Expression = { $_.TimeText } })

                Write-Host ''
                Write-Host "  CSVs imported               : $FilesImported" -ForegroundColor Green
                Write-Host "  Duplicate point exports     : $DuplicatePointExports" -ForegroundColor Yellow
                Write-Host "  Duplicate rows within CSVs  : $WithinFileDuplicates" -ForegroundColor Yellow
                Write-Host "  Overlapping point/timestamps: $OverlappingSamples" -ForegroundColor Yellow
                Write-Host "  Missing values filled       : $BlankValuesFilled" -ForegroundColor Yellow
                Write-Host "  Conflicting values resolved : $Conflicts" -ForegroundColor Yellow
                Write-Host "  Combined timestamps         : $($SortedRows.Count)" -ForegroundColor Green
                Write-Host "  Combined BACnet points      : $($TrendNames.Count)" -ForegroundColor Green
                if ($ImportFailures -gt 0) {
                    Write-Host "  Failed CSV imports          : $ImportFailures" -ForegroundColor Yellow
                }

                if ($SortedRows.Count -gt 1048575) {
                    Write-Host 'Too many data rows for an Excel worksheet. Device skipped.' -ForegroundColor Red
                    $DevicesSkipped++
                    continue
                }
                if ($TrendNames.Count -gt 16383) {
                    Write-Host 'Too many point columns for an Excel worksheet. Device skipped.' -ForegroundColor Red
                    $DevicesSkipped++
                    continue
                }

                $Workbook = $null
                $Worksheet = $null
                $DataRange = $null
                try {
                    $Workbook = $Excel.Workbooks.Add()
                    $Worksheet = $Workbook.Worksheets.Item(1)
                    if (-not $DeviceUnitLabel) { $DeviceUnitLabel = "Dev $DeviceInstance" }
                    $Worksheet.Name = Get-SafeSheetName $DeviceUnitLabel
                    while ($Workbook.Worksheets.Count -gt 1) {
                        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
                    }

                    $TotalDataRows = $SortedRows.Count
                    $TotalColumns = $TrendNames.Count + 1
                    # Write all headers in a single COM call.
                    $HeaderArray = New-Object 'object[,]' 1,$TotalColumns
                    $HeaderArray.SetValue('Time',0,0)
                    $ColumnByTrend = @{}
                    for ($j = 0; $j -lt $TrendNames.Count; $j++) {
                        $HeaderArray.SetValue([string]$TrendNames[$j],0,($j + 1))
                        $ColumnByTrend[[string]$TrendNames[$j]] = $j + 1
                    }
                    $Worksheet.Range(
                        $Worksheet.Cells.Item(1,1),
                        $Worksheet.Cells.Item(1,$TotalColumns)
                    ).Value2 = $HeaderArray

                    if ($TotalDataRows -gt 0) {
                        $DataArray = New-Object 'object[,]' $TotalDataRows,$TotalColumns
                        $culture = [System.Globalization.CultureInfo]::InvariantCulture
                        for ($i = 0; $i -lt $TotalDataRows; $i++) {
                            $current = $SortedRows[$i]
                            if ($null -ne $current.ParsedTime) {
                                $DataArray.SetValue($current.ParsedTime.ToOADate(), $i, 0)
                            }
                            else {
                                $DataArray.SetValue([string]$current.TimeText, $i, 0)
                            }

                            # Only visit populated values, not every point for every timestamp.
                            foreach ($name in $current.Values.Keys) {
                                $columnIndex = [int]$ColumnByTrend[$name]
                                $value = [string]$current.Values[$name]
                                $number = 0.0
                                if ([double]::TryParse($value,
                                        [System.Globalization.NumberStyles]::Any,
                                        $culture,[ref]$number)) {
                                    $DataArray.SetValue($number, $i, $columnIndex)
                                }
                                else {
                                    $DataArray.SetValue($value, $i, $columnIndex)
                                }
                            }
                        }
                        $DataRange = $Worksheet.Range(
                            $Worksheet.Cells.Item(2,1),
                            $Worksheet.Cells.Item($TotalDataRows + 1,$TotalColumns)
                        )
                        $DataRange.Value2 = $DataArray
                        # Show seconds to avoid apparently repeated minute values.
                        $Worksheet.Range(
                            $Worksheet.Cells.Item(2,1),
                            $Worksheet.Cells.Item($TotalDataRows + 1,1)
                        ).NumberFormat = 'm/d/yyyy h:mm:ss'
                    }

                    $header = $Worksheet.Range($Worksheet.Cells.Item(1,1),$Worksheet.Cells.Item(1,$TotalColumns))
                    $header.Font.Bold = $true
                    $filter = $Worksheet.Range($Worksheet.Cells.Item(1,1),$Worksheet.Cells.Item($TotalDataRows + 1,$TotalColumns))
                    $filter.AutoFilter() | Out-Null
                    $Worksheet.Activate()
                    $Excel.ActiveWindow.SplitRow = 1
                    $Excel.ActiveWindow.FreezePanes = $true
                    # AutoFit on hundreds of thousands of cells is extremely slow.
                    # Fixed readable widths avoid full-sheet scanning and COM per-column calls.
                    $Worksheet.Columns.Item(1).ColumnWidth = 22
                    if ($TotalColumns -gt 1) {
                        $Worksheet.Range(
                            $Worksheet.Cells.Item(1,2),
                            $Worksheet.Cells.Item(1,$TotalColumns)
                        ).EntireColumn.ColumnWidth = 24
                    }

                    $FileUnitLabel = ($DeviceUnitLabel -replace '[<>:"/\\|?*]', '_').Trim()
                    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                    $OutputFileName = "Dev $DeviceInstance $FileUnitLabel ALL TRENDLOGS $Timestamp.xlsx"
                    $OutputPath = Join-Path $TrendlogOutputDirectory $OutputFileName
                    $n = 2
                    while (Test-Path -LiteralPath $OutputPath) {
                        $OutputFileName = "Dev $DeviceInstance $FileUnitLabel ALL TRENDLOGS $Timestamp ($n).xlsx"
                        $OutputPath = Join-Path $TrendlogOutputDirectory $OutputFileName
                        $n++
                    }
                    $Workbook.SaveAs($OutputPath,$xlOpenXMLWorkbook)
                    $DevicesCreated++
                    $CreatedWorkbookPaths.Add($OutputPath)
                    Write-Host "SAVED: $OutputPath" -ForegroundColor Green
                }
                catch {
                    $DevicesSkipped++
                    Write-Host "ERROR creating Device $DeviceInstance workbook: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "  Script line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkRed
                }
                finally {
                    if ($Workbook) { try { $Workbook.Close($false) } catch {} }
                    Release-ComObject $DataRange
                    Release-ComObject $Worksheet
                    Release-ComObject $Workbook
                    $Workbook = $null
                    $Worksheet = $null
                    $DataRange = $null
                }
            }

            if ($CreatedWorkbookPaths.Count -gt 0) {
                Write-Host ''
                Write-Host '============================================================' -ForegroundColor Cyan
                Write-Host ' OPTIONAL MASTER WORKBOOK' -ForegroundColor Cyan
                Write-Host '============================================================' -ForegroundColor Cyan
                $CreateMasterWorkbook = Read-Host 'Consolidate all newly created workbooks into one workbook? (y/N) [Default: N]'
                if ($CreateMasterWorkbook -match '^[Yy]') {
                    $MasterWorkbook = $null
                    $TemporarySheet = $null
                    try {
                        $MasterWorkbook = $Excel.Workbooks.Add()
                        $TemporarySheet = $MasterWorkbook.Worksheets.Item(1)
                        $TemporarySheet.Name = 'TEMP'
                        while ($MasterWorkbook.Worksheets.Count -gt 1) {
                            $MasterWorkbook.Worksheets.Item($MasterWorkbook.Worksheets.Count).Delete()
                        }
                        $Copied = 0
                        foreach ($path in $CreatedWorkbookPaths) {
                            $sourceBook = $null
                            try {
                                if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
                                $sourceBook = $Excel.Workbooks.Open($path)
                                for ($sheetIndex = 1; $sheetIndex -le $sourceBook.Worksheets.Count; $sheetIndex++) {
                                    $sourceSheet = $null
                                    $newSheet = $null
                                    try {
                                        $sourceSheet = $sourceBook.Worksheets.Item($sheetIndex)
                                        $name = [string]$sourceSheet.Name
                                        $uniqueName = Get-UniqueSheetName $MasterWorkbook $name
                                        $last = $MasterWorkbook.Worksheets.Item($MasterWorkbook.Worksheets.Count)
                                        $sourceSheet.Copy([Type]::Missing,$last)
                                        Release-ComObject $last
                                        $newSheet = $MasterWorkbook.Worksheets.Item($MasterWorkbook.Worksheets.Count)
                                        $newSheet.Name = $uniqueName
                                        $Copied++
                                        Write-Host "  Added tab: $uniqueName" -ForegroundColor Gray
                                    }
                                    finally {
                                        Release-ComObject $newSheet
                                        Release-ComObject $sourceSheet
                                    }
                                }
                            }
                            catch {
                                Write-Host "  Could not copy $path : $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                            finally {
                                if ($sourceBook) { try { $sourceBook.Close($false) } catch {} }
                                Release-ComObject $sourceBook
                            }
                        }
                        if ($Copied -gt 0) {
                            $TemporarySheet.Delete()
                            $MasterTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                            $MasterOutputFileName = "ALL DEVICES CONSOLIDATED TRENDLOGS $MasterTimestamp.xlsx"
                            $MasterOutputPath = Join-Path $TrendlogOutputDirectory $MasterOutputFileName
                            $count = 2
                            while (Test-Path -LiteralPath $MasterOutputPath) {
                                $MasterOutputFileName = "ALL DEVICES CONSOLIDATED TRENDLOGS $MasterTimestamp ($count).xlsx"
                                $MasterOutputPath = Join-Path $TrendlogOutputDirectory $MasterOutputFileName
                                $count++
                            }
                            $MasterWorkbook.SaveAs($MasterOutputPath,$xlOpenXMLWorkbook)
                            Write-Host "MASTER WORKBOOK SAVED: $MasterOutputPath" -ForegroundColor Green
                            Write-Host "Worksheets copied: $Copied" -ForegroundColor Green
                        }
                        else { Write-Host 'No worksheets copied; master not saved.' -ForegroundColor Yellow }
                    }
                    catch {
                        Write-Host "ERROR creating master workbook: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    finally {
                        if ($MasterWorkbook) { try { $MasterWorkbook.Close($false) } catch {} }
                        Release-ComObject $TemporarySheet
                        Release-ComObject $MasterWorkbook
                    }
                }
                else { Write-Host 'Master workbook consolidation skipped.' -ForegroundColor DarkGray }
            }

            $elapsed = (Get-Date) - $OverallStart
            Write-Host ''
            Write-Host '============================================================' -ForegroundColor Green
            Write-Host ' COMPLETE' -ForegroundColor Green
            Write-Host '============================================================' -ForegroundColor Green
            Write-Host "Device instances found : $($DeviceGroups.Count)"
            Write-Host "Workbooks created      : $DevicesCreated"
            Write-Host "Devices skipped/errors : $DevicesSkipped"
            Write-Host "Output directory       : $TrendlogOutputDirectory"
            Write-Host ('Elapsed time           : {0:hh\:mm\:ss}' -f $elapsed)
        }
        catch {
            Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Script line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
        }
        finally {
            if ($Excel) { try { $Excel.Quit() } catch {} }
            Release-ComObject $Excel
            $Excel = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    $RunAgain = Read-Host 'Enter anything to run again, or press ENTER to exit'
} while (-not [string]::IsNullOrWhiteSpace($RunAgain))

Write-Host 'Exiting.' -ForegroundColor Cyan
