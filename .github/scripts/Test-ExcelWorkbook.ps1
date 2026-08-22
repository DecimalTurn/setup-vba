param(
    [string]$WorkbookPath = "./tests/out/ExcelWorkbook.xlsm",
    [string]$ExpectedValue = "6"
)

$ErrorActionPreference = "Stop"

$fullPath = (Resolve-Path $WorkbookPath).Path
$outputTxt = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($fullPath), "ExcelWorkbook_LocaleTest.txt")

if (Test-Path $outputTxt) {
    Remove-Item $outputTxt -Force
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1 # msoAutomationSecurityForceDisable: run macros without prompting

$workbook = $null
$exitCode = 0

try {
    Write-Host "Opening workbook: $fullPath"
    $workbook = $excel.Workbooks.Open($fullPath)

    Write-Host "Running TestLocaleFormula macro..."
    $excel.Run("'$($workbook.Name)'!TestLocaleFormula")

    if (-not (Test-Path $outputTxt)) {
        throw "Expected output file was not created: $outputTxt"
    }

    $actualValue = (Get-Content -Path $outputTxt -Raw).Trim()

    if ($actualValue -ne $ExpectedValue) {
        throw "Expected '$ExpectedValue' but got '$actualValue'."
    }

    Write-Host "Test passed: TestLocaleFormula produced the expected value '$actualValue'."
} catch {
    Write-Host "::error::Test failed for $WorkbookPath - $($_.Exception.Message)"
    Write-Host "::error::This may indicate a locale/language mismatch (e.g. FormulaLocal argument separator)."
    $exitCode = 1
} finally {
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if (Test-Path $outputTxt) {
        Remove-Item $outputTxt -Force
    }
}

exit $exitCode
