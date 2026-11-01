# Only called by excel_oracle.py after exact curated-source validation.
param(
    [string]$Request,
    [string]$Result,
    [string]$ProcessReceipt,
    [string]$CleanupReceipt
)
$ErrorActionPreference = 'Stop'

function Write-Receipt([string]$Path, $Value) {
    $temporary = $Path + '.tmp'
    $Value | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

if ($CleanupReceipt) {
    $identity = Get-Content -LiteralPath $CleanupReceipt -Raw | ConvertFrom-Json
    if (-not $identity.owned -or $identity.processId -le 0) { throw 'No owned Excel identity' }
    $ownedProcess = Get-Process -Id $identity.processId -ErrorAction SilentlyContinue
    if ($null -eq $ownedProcess) { Write-Output 'Owned Excel already exited'; exit 0 }
    if ($ownedProcess.ProcessName -ne 'EXCEL' -or $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString() -ne $identity.startTicks) {
        throw 'PID was reused or is not the owned Excel process; no process terminated'
    }
    Stop-Process -InputObject $ownedProcess -Force
    $ownedProcess.WaitForExit(5000) | Out-Null
    if (-not $ownedProcess.HasExited) { throw 'Owned Excel has not exited' }
    Write-Output 'Terminated the owned Excel instance'
    exit 0
}

$requestData = Get-Content -LiteralPath $Request -Raw | ConvertFrom-Json
$case = $requestData.case
$directory = $requestData.directory
$existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel = $null
$owned = $false
$books = New-Object System.Collections.Generic.List[object]
$resultData = @{ status = 'failed'; observations = @{}; arrayShapes = @{} }
$exitCode = 1

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OracleExcelWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

function Read-Cell($Cell) {
    $value = $Cell.Value2
    $formula = $null
    if ($Cell.HasFormula) {
        try { $formula = [string]$Cell.Formula2 } catch { $formula = [string]$Cell.Formula }
    }
    # COM error values and literal strings such as '#N/A' must remain distinct.
    if ($excel.WorksheetFunction.IsError($Cell)) {
        $errorType = $null
        try { $errorType = [int]$excel.WorksheetFunction.Error_Type($Cell) } catch {}
        $errors = @{ 1 = '#NULL!'; 2 = '#DIV/0!'; 3 = '#VALUE!'; 4 = '#REF!'; 5 = '#NAME?'; 6 = '#NUM!'; 7 = '#N/A'; 8 = '#GETTING_DATA' }
        if ($null -ne $errorType -and $errors.ContainsKey($errorType)) {
            return @{ value = @{ kind = 'error'; value = $errors[$errorType] }; formula = $formula; errorType = $errorType }
        }
        return @{ value = @{ kind = 'uncomparable'; value = [string]$Cell.Text }; formula = $formula; errorType = $errorType }
    }
    if ($null -eq $value) { $kind = 'blank' }
    elseif ($value -is [bool]) { $kind = 'boolean' }
    elseif ($value -is [string]) { $kind = 'text' }
    elseif ($value -is [ValueType]) { $kind = 'number' }
    else { $kind = 'uncomparable'; $value = [string]$value }
    return @{ value = @{ kind = $kind; value = $value }; formula = $formula }
}

try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$oracleProcessId = 0
    [OracleExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$oracleProcessId) | Out-Null
    if ($oracleProcessId -eq 0 -or $existingExcelIds -contains [int]$oracleProcessId) {
        throw 'COM activation did not provide a new identifiable Excel process; existing instance left untouched'
    }
    $oracleProcess = Get-Process -Id $oracleProcessId
    $owned = $true
    Write-Receipt $ProcessReceipt @{ owned = $true; processId = [int]$oracleProcessId; startTicks = $oracleProcess.StartTime.ToUniversalTime().Ticks.ToString(); hwnd = [long]$excel.Hwnd }

    # ForceDisable must precede EVERY Open. This does not disable XLM, which is
    # why the Python controller accepts only regenerated, macro-free fixtures.
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.Visible = $false
    $resultData.excel = @{ version = [string]$excel.Version; build = [string]$excel.Build; operatingSystem = [string]$excel.OperatingSystem; automationSecurity = [int]$excel.AutomationSecurity }

    foreach ($reference in $case.references) {
        $referencePath = Join-Path $directory $reference
        $referenceBook = $excel.Workbooks.Open($referencePath, 0, $false)
        $books.Add($referenceBook)
    }
    $workbookPath = Join-Path $directory $case.file
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)
    $books.Add($workbook)
    $excel.Calculation = -4135 # xlCalculationManual, set after a workbook exists.
    $excel.Iteration = [bool]$case.iterate
    $excel.MaxIterations = [int]$case.maxIterations
    $excel.MaxChange = [double]$case.maxChange
    $excel.CalculateBeforeSave = $false
    foreach ($book in $books) {
        if ($book.Connections.Count -ne 0) { throw 'Unexpected workbook connection' }
        if ($book.HasVBProject) { throw 'Unexpected VBA project' }
    }
    foreach ($action in $case.actions) {
        $sheet = $workbook.Worksheets.Item($action.sheet)
        if ($action.kind -eq 'formula2') {
            $sheet.Range($action.cell).Formula2 = $action.formula
        } elseif ($action.kind -eq 'filter') {
            $sheet.Range($action.range).AutoFilter([int]$action.field, $action.criteria) | Out-Null
        } else { throw 'Unknown curated action' }
    }
    $resultData.calculation = @{ mode = [int]$excel.Calculation; iteration = [bool]$excel.Iteration; maxIterations = [int]$excel.MaxIterations; maxChange = [double]$excel.MaxChange; date1904 = [bool]$workbook.Date1904; calculateBeforeSave = [bool]$excel.CalculateBeforeSave }
    $excel.CalculateFullRebuild()
    if ($excel.CalculationState -ne 0) { throw 'Excel calculation did not finish' }
    # Saving in manual mode cannot recalculate volatile/circular cells again.
    foreach ($book in $books) { $book.Save() }
    $resultData.calculatedAt = [DateTime]::UtcNow.ToString('o')
    foreach ($selection in $case.targets) {
        $sheet = $workbook.Worksheets.Item($selection.sheet)
        $range = $sheet.Range($selection.range)
        if ($selection.shape) {
            # Formula2 alone is not evidence that this Excel build supports
            # SEQUENCE. Require the native spill extent before using its cells
            # as an array oracle; older builds remain explicitly incomplete.
            try { $spill = $range.Cells.Item(1, 1).SpillingToRange }
            catch { throw 'Native Formula2 spill extent is unavailable on this Excel build' }
            if ($spill.Rows.Count -ne $selection.shape[0] -or $spill.Columns.Count -ne $selection.shape[1]) {
                throw 'Native Formula2 spill does not match the declared fixture shape'
            }
            $resultData.arrayShapes[$selection.nodeId] = @{
                rows = [int]$spill.Rows.Count; columns = [int]$spill.Columns.Count;
                address = [string]$spill.Address()
            }
        }
        if ($range.Rows.Count -eq 1 -and $range.Columns.Count -eq 1) {
            $observation = Read-Cell $range
        } else {
            $matrix = New-Object System.Collections.Generic.List[object]
            for ($rowIndex = 1; $rowIndex -le $range.Rows.Count; $rowIndex++) {
                $rowValues = New-Object System.Collections.Generic.List[object]
                for ($columnIndex = 1; $columnIndex -le $range.Columns.Count; $columnIndex++) {
                    $cellObservation = Read-Cell $range.Cells.Item($rowIndex, $columnIndex)
                    $rowValues.Add($cellObservation.value)
                }
                $matrix.Add($rowValues.ToArray())
            }
            $observation = @{ value = @{ kind = 'array'; rows = [int]$range.Rows.Count; columns = [int]$range.Columns.Count; values = $matrix.ToArray() }; formula = [string]$range.Cells.Item(1, 1).Formula2 }
        }
        $resultData.observations[$selection.nodeId] = $observation
    }
    $resultData.status = 'completed'
    $exitCode = 0
} catch {
    $resultData.error = $_.Exception.Message
} finally {
    Write-Receipt $Result $resultData
    if ($owned) {
        foreach ($book in $books) { try { $book.Close($false) } catch {} }
        try { $excel.Quit() } catch {}
    }
    foreach ($book in $books) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null } catch {} }
    if ($null -ne $excel) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit $exitCode
