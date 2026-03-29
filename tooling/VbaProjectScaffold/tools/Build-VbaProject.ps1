[CmdletBinding()]
param(
    [string]$ConfigPath = '.\config\build-config.json',
    [string]$TemplatePath,
    [string]$SourcePath,
    [string]$OutputPath,
    [string]$OutputFileName,
    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Excel = $null
$script:Workbook = $null
$script:LogEntries = New-Object 'System.Collections.Generic.List[pscustomobject]'
$script:LogLines = New-Object 'System.Collections.Generic.List[string]'
$script:CleanupActions = New-Object 'System.Collections.Generic.List[scriptblock]'
$script:BuildContext = [ordered]@{
    StartTime       = Get-Date
    EndTime         = $null
    DurationSeconds = 0
    OutputDirectory = $null
    OutputFile      = $null
    Template        = $null
    Source          = $null
    ModulesImported = 0
    ClassesImported = 0
    FormsImported   = 0
    References      = @()
    VersionInfo     = $null
    Status          = 'PENDING'
    ConfigPath      = $ConfigPath
}

function Add-CleanupAction {
    param([ScriptBlock]$Action)
    if ($Action) { [void]$script:CleanupActions.Add($Action) }
}

function Invoke-Cleanup {
    for ($i = $script:CleanupActions.Count - 1; $i -ge 0; $i--) {
        try { & $script:CleanupActions[$i] } catch {}
    }
    $script:CleanupActions.Clear()
}

function Write-BuildLog {
    param(
        [string]$Message,
        [string]$Level = 'INFO',
        [string]$Stage
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $entry = [pscustomobject]@{
        timestamp = $timestamp
        level     = $Level
        stage     = $Stage
        message   = $Message
    }
    $script:LogEntries.Add($entry)
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    $script:LogLines.Add($line)
    $color = switch ($Level.ToUpperInvariant()) {
        'ERROR'   { 'Red' }
        'WARNING' { 'Yellow' }
        'SUCCESS' { 'Green' }
        default   { 'White' }
    }
    Write-Host $line -ForegroundColor $color
}

function Close-Excel {
    try {
        if ($script:Workbook) {
            try { $script:Workbook.Close($false) } catch {}
            $script:Workbook = $null
        }
    } finally {
        if ($script:Excel) {
            try {
                $script:Excel.DisplayAlerts = $false
                $script:Excel.Quit()
            } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Excel)
            $script:Excel = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Load-BuildConfig {
    param([string]$Path)
    if (-not $Path) { return $null }
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-BuildLog "Config file not found: $Path" 'WARNING' 'CONFIG'
        return $null
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    } catch {
        $raw = Get-Content -LiteralPath $Path -Raw
    }
    try {
        return $raw | ConvertFrom-Json
    } catch {
        Write-BuildLog "Failed to parse config: $($_.Exception.Message)" 'ERROR' 'CONFIG'
        throw
    }
}

function Resolve-PathSafe {
    param([string]$Path)
    if (-not $Path) { return $null }
    $resolved = Resolve-Path -Path $Path -ErrorAction SilentlyContinue
    if ($resolved) { return $resolved.Path }
    return (Join-Path (Get-Location) $Path)
}

function Reset-VBProjectModules {
    param($VBProject)
    foreach ($component in @($VBProject.VBComponents)) {
        $type = $component.Type
        if ($type -eq 100) { continue }
        try {
            $VBProject.VBComponents.Remove($component)
            Write-BuildLog "Removed module: $($component.Name)" 'INFO' 'VBA'
        } catch {
            Write-BuildLog "Failed to remove module: $($_.Exception.Message)" 'WARNING' 'VBA'
        }
    }
}

function Import-VBAModules {
    param(
        $VBProject,
        [string]$SourcePath
    )

    Write-BuildLog "Importing VBA modules from $SourcePath" 'INFO' 'VBA'

    $summary = [ordered]@{
        Modules = 0
        Classes = 0
        Forms   = 0
    }

    $patterns = @(
        @{ Filter = '*.bas'; Label = 'Module'; Key = 'Modules' },
        @{ Filter = '*.cls'; Label = 'Class';  Key = 'Classes' },
        @{ Filter = '*.frm'; Label = 'Form';   Key = 'Forms' }
    )

    foreach ($pattern in $patterns) {
        $files = Get-ChildItem -Path $SourcePath -Filter $pattern.Filter -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                $VBProject.VBComponents.Import($file.FullName) | Out-Null
                $summary[$pattern.Key]++
                Write-BuildLog "  [OK] $($pattern.Label): $($file.Name)" 'INFO' 'VBA'
            } catch {
                Write-BuildLog "  [FAIL] $($pattern.Label): $($file.Name) :: $($_.Exception.Message)" 'ERROR' 'VBA'
                throw
            }
        }
    }

    Write-BuildLog "Import summary - Modules=$($summary.Modules) Classes=$($summary.Classes) Forms=$($summary.Forms)" 'INFO' 'VBA'
    return [pscustomobject]$summary
}

function Add-References {
    param($VBProject, $ReferenceConfigs)

    Write-BuildLog "Verifying references" 'INFO' 'REFERENCES'

    $resolved = @()
    foreach ($ref in $ReferenceConfigs) {
        $name  = $ref.Name
        $guid  = $ref.GUID
        $major = $ref.Major
        $minor = $ref.Minor
        $exists = $false
        foreach ($existing in $VBProject.References) {
            if ($existing.GUID -eq $guid) { $exists = $true; break }
        }
        if ($exists) {
            Write-BuildLog "  [PRESENT] $name" 'INFO' 'REFERENCES'
        } else {
            $attempt = 0
            $added = $false
            do {
                $attempt++
                try {
                    $VBProject.References.AddFromGuid($guid, $major, $minor) | Out-Null
                    Write-BuildLog "  [ADDED] $name (attempt $attempt)" 'INFO' 'REFERENCES'
                    $added = $true
                } catch {
                    Write-BuildLog "  [WARN] Unable to add $name (attempt $attempt): $($_.Exception.Message)" 'WARNING' 'REFERENCES'
                    Start-Sleep -Milliseconds 300
                }
            } while (-not $added -and $attempt -lt 3)
            if (-not $added) {
                Write-BuildLog "  [FAIL] Reference could not be added: $name" 'ERROR' 'REFERENCES'
            }
        }
        $resolved += [pscustomobject]@{ Name = $name; GUID = $guid; Major = $major; Minor = $minor; Exists = $exists }
    }
    return $resolved
}

function Initialize-WorksheetLayout {
    param(
        $Worksheet,
        $SheetConfig
    )

    $name = $Worksheet.Name
    switch ($name) {
        '入力' {
            $headers = @('ID', '顧客名', 'カテゴリ', '数量', '単価', '金額', '更新日')
            if (-not ($Worksheet.Cells(1,1).Value2)) {
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    $Worksheet.Cells(1, $i + 1).Value2 = $headers[$i]
                    $Worksheet.Cells(1, $i + 1).Interior.Color = 15773696
                    $Worksheet.Cells(1, $i + 1).Font.Bold = $true
                }
                $Worksheet.Columns('A:G').AutoFit()
                Write-BuildLog "Initialized header row on '入力'" 'INFO' 'WORKBOOK'
            }
        }
        '集計' {
            if (-not ($Worksheet.Cells(1,1).Value2)) {
                $Worksheet.Cells(1,1).Value2 = '集計結果'
                $Worksheet.Cells(2,1).Value2 = 'このシートには集計結果を配置します。'
                $Worksheet.Cells(2,1).Font.Italic = $true
                Write-BuildLog "Initialized helper text on '集計'" 'INFO' 'WORKBOOK'
            }
        }
        'マスタ' {
            if (-not ($Worksheet.Cells(1,1).Value2)) {
                $Worksheet.Cells(1,1).Value2 = 'コード'
                $Worksheet.Cells(1,2).Value2 = '名称'
                $Worksheet.Cells(1,3).Value2 = '備考'
                $Worksheet.Columns('A:C').AutoFit()
                Write-BuildLog "Initialized columns on 'マスタ'" 'INFO' 'WORKBOOK'
            }
        }
    }

    if ($SheetConfig.TabColor) {
        try { $Worksheet.Tab.Color = $SheetConfig.TabColor } catch {}
    }
}

function Ensure-WorkbookStructure {
    param($Workbook, $Config)

    $sheetConfigs = @()
    if ($Config -and $Config.Workbook -and $Config.Workbook.Sheets) {
        $sheetConfigs = $Config.Workbook.Sheets
    } else {
        $sheetConfigs = @(
            @{ Name = '入力'; TabColor = '#5B9BD5' },
            @{ Name = '集計'; TabColor = '#A5A5A5' },
            @{ Name = 'マスタ'; TabColor = '#70AD47' },
            @{ Name = 'Logs';  TabColor = '#BFBFBF'; VeryHidden = $true }
        )
    }

    foreach ($sheetConfig in $sheetConfigs) {
        $sheet = $null
        try { $sheet = $Workbook.Worksheets.Item($sheetConfig.Name) } catch {}
        if (-not $sheet) {
            $sheet = $Workbook.Worksheets.Add()
            $sheet.Name = $sheetConfig.Name
            Write-BuildLog "Created worksheet: $($sheetConfig.Name)" 'INFO' 'WORKBOOK'
        } else {
            Write-BuildLog "Worksheet detected: $($sheetConfig.Name)" 'INFO' 'WORKBOOK'
        }

        if ($sheetConfig.VeryHidden) {
            try { $sheet.Visible = 2 } catch {}
        }

        Initialize-WorksheetLayout -Worksheet $sheet -SheetConfig $sheetConfig
    }
}

function Update-VersionInfo {
    param(
        $Workbook,
        $Config,
        $BuildContext
    )

    if (-not $Workbook) { return $null }

    $vi = $null
    if ($Config -and $Config.VersionInfo) { $vi = $Config.VersionInfo }
    if (-not $vi) { return $null }

    $sheetName = if ($vi.Sheet) { [string]$vi.Sheet } else { '集計' }
    try {
        $sheet = $Workbook.Worksheets.Item($sheetName)
    } catch {
        Write-BuildLog "Version info sheet not found: $sheetName" 'WARNING' 'WORKBOOK'
        return $null
    }

    $labelColumn = if ($vi.LabelColumn) { [string]$vi.LabelColumn } else { 'E' }
    $valueColumn = if ($vi.ValueColumn) { [string]$vi.ValueColumn } else { 'F' }

    $versionLabel = if ($vi.VersionLabel) { [string]$vi.VersionLabel } else { 'Version' }
    $releaseLabel = if ($vi.ReleaseDateLabel) { [string]$vi.ReleaseDateLabel } else { 'Release Date' }
    $ownerLabel = if ($vi.OwnerLabel) { [string]$vi.OwnerLabel } else { 'Owner' }
    $changeLabel = if ($vi.ChangeSummaryLabel) { [string]$vi.ChangeSummaryLabel } else { 'Change Summary' }

    $versionValue = if ($vi.VersionNumber) { [string]$vi.VersionNumber } elseif ($env:BUILD_VERSION) { $env:BUILD_VERSION } else { 'v' + (Get-Date -Format 'yyyyMMdd-HHmm') }
    $releaseDate = Get-Date
    $ownerValue = if ($vi.Owner) { [string]$vi.Owner } elseif ($vi.DefaultOwner) { [string]$vi.DefaultOwner } elseif ($env:BUILD_OWNER) { $env:BUILD_OWNER } else { [System.Environment]::UserName }
    $changeSummary = if ($vi.ChangeSummary) { [string]$vi.ChangeSummary } elseif ($env:BUILD_CHANGE_SUMMARY) { $env:BUILD_CHANGE_SUMMARY } elseif ($vi.DefaultChangeSummary) { [string]$vi.DefaultChangeSummary } else { '自動ビルド' }
    $changeSummary = $changeSummary.Trim()
    if ($changeSummary.Length -gt 20) { $changeSummary = $changeSummary.Substring(0, 20) }
    $changeEntry = "{0} {1}" -f $releaseDate.ToString('yyyy/MM/dd'), $changeSummary

    $sheet.Range(("{0}{1}" -f $labelColumn, 2)).Value2 = $versionLabel
    $sheet.Range(("{0}{1}" -f $labelColumn, 2)).Font.Bold = $true
    $sheet.Range(("{0}{1}" -f $valueColumn, 2)).Value2 = $versionValue
    $sheet.Range(("{0}{1}" -f $valueColumn, 2)).NumberFormatLocal = '@'

    $sheet.Range(("{0}{1}" -f $labelColumn, 3)).Value2 = $releaseLabel
    $sheet.Range(("{0}{1}" -f $labelColumn, 3)).Font.Bold = $true
    $sheet.Range(("{0}{1}" -f $valueColumn, 3)).Value2 = $releaseDate
    $sheet.Range(("{0}{1}" -f $valueColumn, 3)).NumberFormatLocal = 'yyyy/m/d h:mm'

    $sheet.Range(("{0}{1}" -f $labelColumn, 4)).Value2 = $ownerLabel
    $sheet.Range(("{0}{1}" -f $labelColumn, 4)).Font.Bold = $true
    $sheet.Range(("{0}{1}" -f $valueColumn, 4)).Value2 = $ownerValue
    $sheet.Range(("{0}{1}" -f $valueColumn, 4)).NumberFormatLocal = '@'

    $sheet.Range(("{0}{1}" -f $labelColumn, 5)).Value2 = $changeLabel
    $sheet.Range(("{0}{1}" -f $labelColumn, 5)).Font.Bold = $true

    $changeRangeAddress = if ($vi.ChangeSummaryRange) { [string]$vi.ChangeSummaryRange } else { '{0}{1}:{0}{2}' -f $valueColumn, 6, 10 }
    $changeRange = $sheet.Range($changeRangeAddress)
    $existing = @()
    foreach ($cell in $changeRange.Cells) {
        $value = ("" + $cell.Value2).Trim()
        if ($value.Length -gt 0) { $existing += $value }
    }
    $updated = @($changeEntry) + $existing
    for ($i = 0; $i -lt $changeRange.Cells.Count; $i++) {
        $cell = $changeRange.Cells.Item($i + 1)
        if ($i -lt $updated.Count) {
            $cell.Value2 = $updated[$i]
            $cell.NumberFormatLocal = '@'
        } else {
            $cell.ClearContents()
        }
    }

    $sheet.Columns("$labelColumn:$valueColumn").AutoFit()

    return [pscustomobject]@{
        Sheet = $sheetName
        Version = $versionValue
        ReleaseDate = $releaseDate
        Owner = $ownerValue
        ChangeEntry = $changeEntry
        ChangeRange = $changeRangeAddress
    }
}

function Save-Logs {
    param(
        [string]$Directory,
        [string]$BaseName,
        [pscustomobject]$Summary
    )

    if (-not $Directory) { $Directory = '.\\logs' }
    if (-not (Test-Path -LiteralPath $Directory)) {
        New-Item -ItemType Directory -Path $Directory -Force | Out-Null
    }

    $textPath = Join-Path $Directory ("$BaseName.log")
    $jsonPath = Join-Path $Directory ("$BaseName.json")
    $summaryPath = Join-Path $Directory ("$BaseName-summary.json")

    $script:LogLines | Set-Content -LiteralPath $textPath -Encoding UTF8
    $script:LogEntries | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    $Summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

    return [pscustomobject]@{
        TextLog = $textPath
        JsonLog = $jsonPath
        Summary = $summaryPath
    }
}

$exitCode = 0
$logBaseName = 'build-' + (Get-Date -Format 'yyyyMMdd_HHmmss')
$script:BuildContext.LogBaseName = $logBaseName

try {
    $config = Load-BuildConfig -Path $ConfigPath

    if (-not $TemplatePath) { $TemplatePath = $config.TemplatePath }
    if (-not $SourcePath)   { $SourcePath   = $config.SourcePath }
    if (-not $OutputPath)   { $OutputPath   = $config.OutputPath }
    if (-not $OutputFileName) { $OutputFileName = $config.OutputFileName }
    if (-not $LogDirectory) { $LogDirectory = if ($config.LogDirectory) { $config.LogDirectory } else { '.\\logs' } }

    if (-not $TemplatePath) { $TemplatePath = '.\\template\\BaseTemplate.xltm' }
    if (-not $SourcePath)   { $SourcePath   = '.\\src\\vba' }
    if (-not $OutputPath)   { $OutputPath   = '.\\build' }
    if (-not $OutputFileName) { $OutputFileName = 'Output_#{timestamp}.xlsm' }

    $script:BuildContext.Template = Resolve-PathSafe $TemplatePath
    $script:BuildContext.Source = Resolve-PathSafe $SourcePath
    $script:BuildContext.OutputDirectory = Resolve-PathSafe $OutputPath

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        Write-BuildLog "Template not found: $TemplatePath" 'ERROR' 'VALIDATION'
        throw 'TemplateNotFound'
    }
    if (-not (Test-Path -LiteralPath $SourcePath)) {
        Write-BuildLog "Source folder not found: $SourcePath" 'ERROR' 'VALIDATION'
        throw 'SourceNotFound'
    }
    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-BuildLog "Created output folder: $OutputPath" 'INFO' 'VALIDATION'
    }
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        Write-BuildLog "Created log folder: $LogDirectory" 'INFO' 'VALIDATION'
    }

    $timestampToken = Get-Date -Format 'yyyyMMdd_HHmmss'
    if ($OutputFileName -match '#\{timestamp\}') {
        $OutputFileName = $OutputFileName -replace '#\{timestamp\}', $timestampToken
    }
    if ($OutputFileName -match '#\{date\}') {
        $OutputFileName = $OutputFileName -replace '#\{date\}', (Get-Date -Format 'yyyyMMdd')
    }

    $outputFullPath = Join-Path $OutputPath $OutputFileName
    $script:BuildContext.OutputFile = $outputFullPath

    Write-BuildLog 'Starting Excel' 'INFO' 'EXCEL'
    $maxAttempts = 3
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $script:Excel = New-Object -ComObject Excel.Application
            break
        } catch {
            Write-BuildLog "Excel launch failed (attempt $attempt/$maxAttempts): $($_.Exception.Message)" 'WARNING' 'EXCEL'
            Start-Sleep -Seconds 1
        }
    }
    if (-not $script:Excel) {
        Write-BuildLog 'Excel could not be started' 'ERROR' 'EXCEL'
        throw 'ExcelLaunchFailed'
    }

    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false

    $templateFullPath = Resolve-PathSafe $TemplatePath
    Write-BuildLog "Opening template: $templateFullPath" 'INFO' 'EXCEL'
    $script:Workbook = $script:Excel.Workbooks.Add($templateFullPath)

    try {
        $vbProject = $script:Workbook.VBProject
    } catch {
        Write-BuildLog 'VBA project is not accessible. Enable programmatic access to the VBA project.' 'ERROR' 'EXCEL'
        throw
    }

    Reset-VBProjectModules -VBProject $vbProject
    $importSummary = Import-VBAModules -VBProject $vbProject -SourcePath $SourcePath
    $script:BuildContext.ModulesImported = $importSummary.Modules
    $script:BuildContext.ClassesImported = $importSummary.Classes
    $script:BuildContext.FormsImported = $importSummary.Forms

    $references = if ($config -and $config.References) { $config.References } else {
        @(
            @{ Name = 'Microsoft Scripting Runtime'; GUID = '{420B2830-E718-11CF-893D-00A0C9054228}'; Major = 1; Minor = 0 },
            @{ Name = 'Microsoft ActiveX Data Objects 6.1 Library'; GUID = '{B691E011-1797-432E-907A-4D8C69339129}'; Major = 6; Minor = 1 }
        )
    }
    $referenceSummary = Add-References -VBProject $vbProject -ReferenceConfigs $references
    $script:BuildContext.References = $referenceSummary

    Ensure-WorkbookStructure -Workbook $script:Workbook -Config $config
    $versionInfo = Update-VersionInfo -Workbook $script:Workbook -Config $config -BuildContext $script:BuildContext
    $script:BuildContext.VersionInfo = $versionInfo

    Write-BuildLog "Saving workbook: $outputFullPath" 'INFO' 'SAVE'
    $saveSucceeded = $false
    try {
        $script:Workbook.SaveAs($outputFullPath, 52)
        $saveSucceeded = $true
    } catch {
        Write-BuildLog "Save failed: $($_.Exception.Message)" 'WARNING' 'SAVE'
        $retryPath = Join-Path $OutputPath ('Retry_' + $OutputFileName)
        try {
            $script:Workbook.SaveAs($retryPath, 52)
            $outputFullPath = $retryPath
            $script:BuildContext.OutputFile = $retryPath
            $saveSucceeded = $true
            Write-BuildLog "Retry save succeeded: $retryPath" 'INFO' 'SAVE'
        } catch {
            Write-BuildLog "Retry save failed: $($_.Exception.Message)" 'ERROR' 'SAVE'
        }
    }

    if (-not $saveSucceeded) {
        throw 'SaveFailed'
    }

    $script:BuildContext.Status = 'SUCCESS'
    Write-BuildLog "Build completed successfully: $outputFullPath" 'SUCCESS' 'SUMMARY'
}
catch {
    $exitCode = 1
    $script:BuildContext.Status = 'FAILED'
    Write-BuildLog "Fatal error: $($_.Exception.Message)" 'ERROR' 'SUMMARY'
    if ($script:BuildContext.OutputFile -and (Test-Path -LiteralPath $script:BuildContext.OutputFile)) {
        try {
            Remove-Item -LiteralPath $script:BuildContext.OutputFile -Force
            Write-BuildLog 'Removed incomplete output file' 'INFO' 'CLEANUP'
        } catch {}
    }
}
finally {
    $script:BuildContext.EndTime = Get-Date
    $script:BuildContext.DurationSeconds = [Math]::Round(($script:BuildContext.EndTime - $script:BuildContext.StartTime).TotalSeconds, 2)
    $summaryObject = [pscustomobject]$script:BuildContext
    $logInfo = Save-Logs -Directory $LogDirectory -BaseName $logBaseName -Summary $summaryObject
    Write-BuildLog "Log file: $($logInfo.TextLog)" 'INFO' 'SUMMARY'
    Write-BuildLog "Log JSON: $($logInfo.JsonLog)" 'INFO' 'SUMMARY'
    Close-Excel
    Invoke-Cleanup
}

exit $exitCode

