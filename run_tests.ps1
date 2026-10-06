[Reflection.Assembly]::LoadFrom("e:\Excel\bin\EPPlus.dll") | Out-Null

$baseUrl = "http://localhost:51234/default.aspx"
$LF = "`r`n"

Write-Host "=========================================================="
Write-Host "TEST 1: Inspect test_gpfpran.xlsx (Auto-Detection)"
Write-Host "=========================================================="
$fileBytes1 = [System.IO.File]::ReadAllBytes("e:\Excel\test_gpfpran.xlsx")
$boundary1 = [System.Guid]::NewGuid().ToString()

$bodyParts1 = (
    "--$boundary1",
    "Content-Disposition: form-data; name=`"excelFile`"; filename=`"test_gpfpran.xlsx`"",
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$LF",
    [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($fileBytes1),
    "--$boundary1--$LF"
) -join $LF

$inspectRes1 = Invoke-RestMethod -Uri "$($baseUrl)?action=inspect" -Method Post `
    -ContentType "multipart/form-data; boundary=$boundary1" `
    -Body ([System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($bodyParts1))

Write-Host "Detected Key Type: $($inspectRes1.detectedKeyType)"
Write-Host "Detected Column:   $($inspectRes1.detectedColumn)"

if ($inspectRes1.detectedKeyType -eq "GPFPRAN" -and $inspectRes1.detectedColumn -eq "GpfPran") {
    Write-Host "-> PASS: Auto-detected GpfPran correctly!" -ForegroundColor Green
} else {
    Write-Host "-> FAIL: Expected GPFPRAN" -ForegroundColor Red
}

Write-Host "`n=========================================================="
Write-Host "TEST 2: Process test_gpfpran.xlsx (Inject PCNO, PIS, and Bank ACCNO)"
Write-Host "=========================================================="
$boundary2 = [System.Guid]::NewGuid().ToString()
$bodyPartsProc1 = (
    "--$boundary2",
    "Content-Disposition: form-data; name=`"excelFile`"; filename=`"test_gpfpran.xlsx`"",
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$LF",
    [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($fileBytes1),
    "--$boundary2",
    "Content-Disposition: form-data; name=`"inputCol`"$LF",
    "GpfPran",
    "--$boundary2",
    "Content-Disposition: form-data; name=`"keyType`"$LF",
    "GPFPRAN",
    "--$boundary2",
    "Content-Disposition: form-data; name=`"outputCols`"$LF",
    "PCNO,PIS,BANK_ACCNO",
    "--$boundary2--$LF"
) -join $LF

Invoke-RestMethod -Uri "$($baseUrl)?action=process" -Method Post `
    -ContentType "multipart/form-data; boundary=$boundary2" `
    -Body ([System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($bodyPartsProc1)) `
    -OutFile "e:\Excel\out_gpfpran.xlsx"

$outPkg1 = New-Object OfficeOpenXml.ExcelPackage([System.IO.FileInfo]::new("e:\Excel\out_gpfpran.xlsx"))
$outWs1 = $outPkg1.Workbook.Worksheets[1]

Write-Host "Headers: Col4 = $($outWs1.Cells[1, 4].Text), Col5 = $($outWs1.Cells[1, 5].Text), Col6 = $($outWs1.Cells[1, 6].Text)"
Write-Host "Row 2: $($outWs1.Cells[2, 2].Text) -> PCNO: $($outWs1.Cells[2, 4].Text), PIS: $($outWs1.Cells[2, 5].Text), Bank A/C: $($outWs1.Cells[2, 6].Text)"
Write-Host "Row 3: $($outWs1.Cells[3, 2].Text) -> PCNO: $($outWs1.Cells[3, 4].Text), PIS: $($outWs1.Cells[3, 5].Text), Bank A/C: $($outWs1.Cells[3, 6].Text)"
Write-Host "Row 4: $($outWs1.Cells[4, 2].Text) -> PCNO: $($outWs1.Cells[4, 4].Text), PIS: $($outWs1.Cells[4, 5].Text), Bank A/C: $($outWs1.Cells[4, 6].Text)"
Write-Host "Row 5: $($outWs1.Cells[5, 2].Text) -> PCNO: $($outWs1.Cells[5, 4].Text), PIS: $($outWs1.Cells[5, 5].Text), Bank A/C: $($outWs1.Cells[5, 6].Text)"

if ($outWs1.Cells[2, 4].Text -eq "5001" -and $outWs1.Cells[2, 5].Text -eq "2008AE10" -and $outWs1.Cells[2, 6].Text -eq "10000000001" -and
    $outWs1.Cells[3, 4].Text -eq "5010" -and $outWs1.Cells[3, 5].Text -eq "2008AE12" -and $outWs1.Cells[3, 6].Text -eq "10000000010" -and
    $outWs1.Cells[5, 4].Text -eq "Not Found") {
    Write-Host "-> PASS: test_gpfpran.xlsx matched and appended correctly with Bank Account!" -ForegroundColor Green
} else {
    Write-Host "-> FAIL: Unexpected values in out_gpfpran.xlsx" -ForegroundColor Red
}
$outPkg1.Dispose()

Write-Host "`n=========================================================="
Write-Host "TEST 3: Process test_pis_max.xlsx (MAX(PCNO) Rule + Bank Account Number)"
Write-Host "=========================================================="
$fileBytes2 = [System.IO.File]::ReadAllBytes("e:\Excel\test_pis_max.xlsx")
$boundary3 = [System.Guid]::NewGuid().ToString()
$bodyPartsProc2 = (
    "--$boundary3",
    "Content-Disposition: form-data; name=`"excelFile`"; filename=`"test_pis_max.xlsx`"",
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$LF",
    [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($fileBytes2),
    "--$boundary3",
    "Content-Disposition: form-data; name=`"inputCol`"$LF",
    "PIS NO",
    "--$boundary3",
    "Content-Disposition: form-data; name=`"keyType`"$LF",
    "PIS",
    "--$boundary3",
    "Content-Disposition: form-data; name=`"outputCols`"$LF",
    "PCNO,BANK_ACCNO,GPFPRAN",
    "--$boundary3--$LF"
) -join $LF

Invoke-RestMethod -Uri "$($baseUrl)?action=process" -Method Post `
    -ContentType "multipart/form-data; boundary=$boundary3" `
    -Body ([System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($bodyPartsProc2)) `
    -OutFile "e:\Excel\out_pis_max.xlsx"

$outPkg2 = New-Object OfficeOpenXml.ExcelPackage([System.IO.FileInfo]::new("e:\Excel\out_pis_max.xlsx"))
$outWs2 = $outPkg2.Workbook.Worksheets[1]

Write-Host "Headers: Col4 = $($outWs2.Cells[1, 4].Text), Col5 = $($outWs2.Cells[1, 5].Text), Col6 = $($outWs2.Cells[1, 6].Text)"
Write-Host "Row 2: PIS $($outWs2.Cells[2, 2].Text) -> PCNO: $($outWs2.Cells[2, 4].Text), Bank A/C: $($outWs2.Cells[2, 5].Text), GPFPRAN: $($outWs2.Cells[2, 6].Text)"
Write-Host "Row 3: PIS $($outWs2.Cells[3, 2].Text) -> PCNO: $($outWs2.Cells[3, 4].Text), Bank A/C: $($outWs2.Cells[3, 5].Text), GPFPRAN: $($outWs2.Cells[3, 6].Text)"
Write-Host "Row 4: PIS $($outWs2.Cells[4, 2].Text) -> PCNO: $($outWs2.Cells[4, 4].Text), Bank A/C: $($outWs2.Cells[4, 5].Text), GPFPRAN: $($outWs2.Cells[4, 6].Text)"
Write-Host "Row 5: PIS $($outWs2.Cells[5, 2].Text) -> PCNO: $($outWs2.Cells[5, 4].Text), Bank A/C: $($outWs2.Cells[5, 5].Text), GPFPRAN: $($outWs2.Cells[5, 6].Text)"

$p2 = $outWs2.Cells[2, 4].Text
$b2 = $outWs2.Cells[2, 5].Text
$g2 = $outWs2.Cells[2, 6].Text

$p3 = $outWs2.Cells[3, 4].Text
$b3 = $outWs2.Cells[3, 5].Text
$g3 = $outWs2.Cells[3, 6].Text

$p4 = $outWs2.Cells[4, 4].Text
$b4 = $outWs2.Cells[4, 5].Text
$g4 = $outWs2.Cells[4, 6].Text

if ($p2 -eq "5001" -and $b2 -eq "10000000001" -and $g2 -eq "GPF-1111" -and
    $p3 -eq "5010" -and $b3 -eq "10000000010" -and $g3 -eq "GPF-2222" -and
    $p4 -eq "5025" -and $b4 -eq "10000000025" -and $g4 -eq "GPF-5555") {
    Write-Host "-> PASS: Correctly resolved! PIS 2008AE12 => PCNO $p3, Bank A/C $b3; PIS 2015EF48 => PCNO $p4, Bank A/C $b4" -ForegroundColor Green
} else {
    Write-Host "-> FAIL: Unexpected values in out_pis_max.xlsx" -ForegroundColor Red
}
$outPkg2.Dispose()

Write-Host "`n=========================================================="
Write-Host "TEST 4: Manual Mapping on test_custom_col.xlsx"
Write-Host "=========================================================="
$fileBytes3 = [System.IO.File]::ReadAllBytes("e:\Excel\test_custom_col.xlsx")
$boundary4 = [System.Guid]::NewGuid().ToString()
$bodyPartsProc3 = (
    "--$boundary4",
    "Content-Disposition: form-data; name=`"excelFile`"; filename=`"test_custom_col.xlsx`"",
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$LF",
    [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($fileBytes3),
    "--$boundary4",
    "Content-Disposition: form-data; name=`"inputCol`"$LF",
    "EmpCode",
    "--$boundary4",
    "Content-Disposition: form-data; name=`"keyType`"$LF",
    "PCNO",
    "--$boundary4",
    "Content-Disposition: form-data; name=`"outputCols`"$LF",
    "PIS,BANK_ACCNO,GPFPRAN",
    "--$boundary4--$LF"
) -join $LF

Invoke-RestMethod -Uri "$($baseUrl)?action=process" -Method Post `
    -ContentType "multipart/form-data; boundary=$boundary4" `
    -Body ([System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($bodyPartsProc3)) `
    -OutFile "e:\Excel\out_custom.xlsx"

$outPkg3 = New-Object OfficeOpenXml.ExcelPackage([System.IO.FileInfo]::new("e:\Excel\out_custom.xlsx"))
$outWs3 = $outPkg3.Workbook.Worksheets[1]

Write-Host "Headers: Col4 = $($outWs3.Cells[1, 4].Text), Col5 = $($outWs3.Cells[1, 5].Text), Col6 = $($outWs3.Cells[1, 6].Text)"
Write-Host "Row 2: EmpCode $($outWs3.Cells[2, 2].Text) -> PIS: $($outWs3.Cells[2, 4].Text), Bank A/C: $($outWs3.Cells[2, 5].Text), GPFPRAN: $($outWs3.Cells[2, 6].Text)"
Write-Host "Row 3: EmpCode $($outWs3.Cells[3, 2].Text) -> PIS: $($outWs3.Cells[3, 4].Text), Bank A/C: $($outWs3.Cells[3, 5].Text), GPFPRAN: $($outWs3.Cells[3, 6].Text)"

if ($outWs3.Cells[2, 4].Text -eq "2008AE10" -and $outWs3.Cells[2, 5].Text -eq "10000000001" -and $outWs3.Cells[2, 6].Text -eq "GPF-1111" -and
    $outWs3.Cells[3, 4].Text -eq "2008AE12" -and $outWs3.Cells[3, 5].Text -eq "10000000010" -and $outWs3.Cells[3, 6].Text -eq "GPF-2222") {
    Write-Host "-> PASS: Manual mapping with PCNO lookup passed flawlessly!" -ForegroundColor Green
} else {
    Write-Host "-> FAIL: Unexpected values in out_custom.xlsx" -ForegroundColor Red
}
$outPkg3.Dispose()

Write-Host "`nAll validation tests complete!"
