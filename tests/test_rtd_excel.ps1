# PowerShell script to test RTD Server in Excel
$ErrorActionPreference = "Stop"

$dllPath = Resolve-Path "..\build\MyHybridServer.xll"
$progID = "My.Hybrid.Server"

Write-Host "1. Registering DLL..."
regsvr32.exe /s $dllPath
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to register DLL."
}
Write-Host "DLL Registered."

Write-Host "2. Starting Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)

    Write-Host "3. Inserting RTD Formula..."
    # Formula: =RTD("My.Hybrid.Server",, "Topic1")
    $cell = $sheet.Range("A1")
    $cell.Formula = "=RTD(`"$progID`",, `"Topic1`")"

    # Wait for RTD to update (it might be async)
    Start-Sleep -Seconds 2

    $value = $cell.Value()
    Write-Host "Cell Value: $value"

    if ($value -eq "Connecting...") {
        Write-Host "SUCCESS: RTD returned expected value." -ForegroundColor Green
    } else {
        Write-Error "FAILURE: Unexpected value returned."
    }

} catch {
    Write-Error "An error occurred: $_"
} finally {
    Write-Host "4. Cleaning up..."
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Unregister
    regsvr32.exe /s /u $dllPath
    Write-Host "Done."
}
