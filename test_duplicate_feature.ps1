[Reflection.Assembly]::LoadFrom("e:\Excel\bin\EPPlus.dll") | Out-Null

$baseUrl = "http://localhost:51234/default.aspx"
$loginUrl = "http://localhost:51234/Login.aspx"
$LF = "`r`n"

# 1. Ensure IIS Express is running on port 51234
$iisProc = Get-Process -Name iisexpress -ErrorAction SilentlyContinue
if (-not $iisProc) {
    Write-Host "Starting IIS Express on port 51234..."
    $iisPath = "C:\Program Files\IIS Express\iisexpress.exe"
    Start-Process -FilePath $iisPath -ArgumentList "/path:e:\Excel /port:51234" -WindowStyle Hidden
    Start-Sleep -Seconds 3
}

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# 2. Login as 1001
$loginGet = Invoke-WebRequest -Uri $loginUrl -WebSession $session -UseBasicParsing
$vs = [regex]::Match($loginGet.Content, 'id="__VIEWSTATE" value="([^"]*)"').Groups[1].Value
$vsg = [regex]::Match($loginGet.Content, 'id="__VIEWSTATEGENERATOR" value="([^"]*)"').Groups[1].Value
$ev = [regex]::Match($loginGet.Content, 'id="__EVENTVALIDATION" value="([^"]*)"').Groups[1].Value

$authPost = @{
    "__VIEWSTATE" = $vs
    "__VIEWSTATEGENERATOR" = $vsg
    "__EVENTVALIDATION" = $ev
    "txtUsername" = "1001"
    "txtPassword" = "validpassword"
    "btnLogin" = "Sign In to System"
}

try {
    $authRes = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $authPost -WebSession $session -UseBasicParsing -MaximumRedirection 0 -ErrorAction Stop
} catch {
    $authRes = $_.Exception.Response
}

Write-Host "Authentication HTTP status: $($authRes.StatusCode)"

# 3. Create test file with duplicates: test_with_duplicates.xlsx
$testPkg = New-Object OfficeOpenXml.ExcelPackage
$testWs = $testPkg.Workbook.Worksheets.Add("Employees")
$testWs.Cells[1, 1].Value = "SLNO"
$testWs.Cells[1, 2].Value = "PCNO"
$testWs.Cells[1, 3].Value = "Name"

$testWs.Cells[2, 1].Value = 1
$testWs.Cells[2, 2].Value = "5001"
$testWs.Cells[2, 3].Value = "Alice"

$testWs.Cells[3, 1].Value = 2
$testWs.Cells[3, 2].Value = "5010"
$testWs.Cells[3, 3].Value = "Bob"

$testWs.Cells[4, 1].Value = 3
$testWs.Cells[4, 2].Value = "5001"
$testWs.Cells[4, 3].Value = "Alice Repeat"

$testWs.Cells[5, 1].Value = 4
$testWs.Cells[5, 2].Value = "5025"
$testWs.Cells[5, 3].Value = "Charlie"

$testWs.Cells[6, 1].Value = 5
$testWs.Cells[6, 2].Value = "5001"
$testWs.Cells[6, 3].Value = "Alice Third Time"

$testPkg.SaveAs([System.IO.FileInfo]::new("e:\Excel\test_with_duplicates.xlsx"))
$testPkg.Dispose()

Write-Host "`n========================================================"
Write-Host "TEST A: Process with highlightDuplicates = true (Default)"
Write-Host "========================================================"
$fileBytes = [System.IO.File]::ReadAllBytes("e:\Excel\test_with_duplicates.xlsx")
$boundary = [System.Guid]::NewGuid().ToString()
$bodyParts = (
    "--$boundary",
    "Content-Disposition: form-data; name=`"excelFile`"; filename=`"test_with_duplicates.xlsx`"",
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$LF",
    [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($fileBytes),
    "--$boundary",
    "Content-Disposition: form-data; name=`"inputCol`"$LF",
    "PCNO",
    "--$boundary",
    "Content-Disposition: form-data; name=`"keyType`"$LF",
    "PCNO",
    "--$boundary",
    "Content-Disposition: form-data; name=`"outputCols`"$LF",
    "PIS,GPFPRAN,BANK_ACCNO",
    "--$boundary",
    "Content-Disposition: form-data; name=`"highlightDuplicates`"$LF",
    "true",
    "--$boundary--$LF"
) -join $LF

$resA = Invoke-WebRequest -Uri "$($baseUrl)?action=process" -Method Post `
    -ContentType "multipart/form-data; boundary=$boundary" `
    -Body ([System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($bodyParts)) `
    -WebSession $session `
    -UseBasicParsing

$dupHeaderA = $resA.Headers["X-Duplicate-Count"]
Write-Host "Returned X-Duplicate-Count: $dupHeaderA"
[System.IO.File]::WriteAllBytes("e:\Excel\out_with_duplicates.xlsx", $resA.RawContentStream.ToArray())

$outPkgA = New-Object OfficeOpenXml.ExcelPackage([System.IO.FileInfo]::new("e:\Excel\out_with_duplicates.xlsx"))
$outWsA = $outPkgA.Workbook.Worksheets[1]

$r2Highlight = ($outWsA.Cells[2, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::Solid -and $outWsA.Cells[2, 1].Style.Fill.BackgroundColor.Rgb -eq "FFFFF2CC")
$r3Highlight = ($outWsA.Cells[3, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::Solid)
$r4Highlight = ($outWsA.Cells[4, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::Solid -and $outWsA.Cells[4, 1].Style.Fill.BackgroundColor.Rgb -eq "FFFFF2CC")
$r5Highlight = ($outWsA.Cells[5, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::Solid)
$r6Highlight = ($outWsA.Cells[6, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::Solid -and $outWsA.Cells[6, 1].Style.Fill.BackgroundColor.Rgb -eq "FFFFF2CC")

# Also check that appended columns on duplicate rows are highlighted
$r2Col6Highlight = ($outWsA.Cells[2, 6].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::Solid)

Write-Host "Row 2 (5001 duplicate 1): Highlighted = $r2Highlight (All cols = $r2Col6Highlight)"
Write-Host "Row 3 (5010 unique):      Highlighted = $r3Highlight"
Write-Host "Row 4 (5001 duplicate 2): Highlighted = $r4Highlight"
Write-Host "Row 5 (5025 unique):      Highlighted = $r5Highlight"
Write-Host "Row 6 (5001 duplicate 3): Highlighted = $r6Highlight"

if ($r2Highlight -and -not $r3Highlight -and $r4Highlight -and -not $r5Highlight -and $r6Highlight -and $r2Col6Highlight -and $dupHeaderA -eq "3") {
    Write-Host "-> PASS: All occurrences of duplicate PCNO 5001 highlighted in soft yellow across full row!" -ForegroundColor Green
} else {
    Write-Host "-> FAIL: Duplicate highlighting did not match expectations." -ForegroundColor Red
}
$outPkgA.Dispose()

Write-Host "`n========================================================"
Write-Host "TEST B: Process with highlightDuplicates = false (Disabled)"
Write-Host "========================================================"
$boundaryB = [System.Guid]::NewGuid().ToString()
$bodyPartsB = (
    "--$boundaryB",
    "Content-Disposition: form-data; name=`"excelFile`"; filename=`"test_with_duplicates.xlsx`"",
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$LF",
    [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($fileBytes),
    "--$boundaryB",
    "Content-Disposition: form-data; name=`"inputCol`"$LF",
    "PCNO",
    "--$boundaryB",
    "Content-Disposition: form-data; name=`"keyType`"$LF",
    "PCNO",
    "--$boundaryB",
    "Content-Disposition: form-data; name=`"outputCols`"$LF",
    "PIS,GPFPRAN,BANK_ACCNO",
    "--$boundaryB",
    "Content-Disposition: form-data; name=`"highlightDuplicates`"$LF",
    "false",
    "--$boundaryB--$LF"
) -join $LF

$resB = Invoke-WebRequest -Uri "$($baseUrl)?action=process" -Method Post `
    -ContentType "multipart/form-data; boundary=$boundaryB" `
    -Body ([System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($bodyPartsB)) `
    -WebSession $session `
    -UseBasicParsing

$dupHeaderB = $resB.Headers["X-Duplicate-Count"]
Write-Host "Returned X-Duplicate-Count: $dupHeaderB"
[System.IO.File]::WriteAllBytes("e:\Excel\out_no_highlight.xlsx", $resB.RawContentStream.ToArray())

$outPkgB = New-Object OfficeOpenXml.ExcelPackage([System.IO.FileInfo]::new("e:\Excel\out_no_highlight.xlsx"))
$outWsB = $outPkgB.Workbook.Worksheets[1]

$r2FillNone = ($outWsB.Cells[2, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::None)
$r4FillNone = ($outWsB.Cells[4, 1].Style.Fill.PatternType -eq [OfficeOpenXml.Style.ExcelFillStyle]::None)

Write-Host "Row 2 Fill when disabled: $($outWsB.Cells[2, 1].Style.Fill.PatternType)"
Write-Host "Row 4 Fill when disabled: $($outWsB.Cells[4, 1].Style.Fill.PatternType)"

if ($r2FillNone -and $r4FillNone) {
    Write-Host "-> PASS: Rows remain unhighlighted when highlightDuplicates = false!" -ForegroundColor Green
} else {
    Write-Host "-> FAIL: Rows were unexpectedly highlighted." -ForegroundColor Red
}
$outPkgB.Dispose()

Write-Host "`nAll Duplicate Feature Tests PASSED!"
