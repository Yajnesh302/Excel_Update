[Reflection.Assembly]::LoadFrom("e:\Excel\bin\EPPlus.dll") | Out-Null

# 1. Create test_gpfpran.xlsx (simulating Temp_Sh_Lpt_Sep)
$pkg1 = New-Object OfficeOpenXml.ExcelPackage
$ws1 = $pkg1.Workbook.Worksheets.Add("Sheet1")
$ws1.Cells[1, 1].Value = "SLNO"
$ws1.Cells[1, 2].Value = "GpfPran"
$ws1.Cells[1, 3].Value = "Remarks"

$ws1.Cells[2, 1].Value = 1
$ws1.Cells[2, 2].Value = "GPF-1111"
$ws1.Cells[2, 3].Value = "Test 1"

$ws1.Cells[3, 1].Value = 2
$ws1.Cells[3, 2].Value = "GPF-2222"
$ws1.Cells[3, 3].Value = "Test 2 (Has 2 PCNOs 5002 and 5010)"

$ws1.Cells[4, 1].Value = 3
$ws1.Cells[4, 2].Value = "PRAN-3333"
$ws1.Cells[4, 3].Value = "Test 3"

$ws1.Cells[5, 1].Value = 4
$ws1.Cells[5, 2].Value = "UNKNOWN-99"
$ws1.Cells[5, 3].Value = "Invalid Record"

$pkg1.SaveAs([System.IO.FileInfo]::new("e:\Excel\test_gpfpran.xlsx"))
$pkg1.Dispose()

# 2. Create test_pis_max.xlsx (testing resolution for PIS like 2008AE10)
$pkg2 = New-Object OfficeOpenXml.ExcelPackage
$ws2 = $pkg2.Workbook.Worksheets.Add("Sheet1")
$ws2.Cells[1, 1].Value = "SLNO"
$ws2.Cells[1, 2].Value = "PIS NO"
$ws2.Cells[1, 3].Value = "Name"

$ws2.Cells[2, 1].Value = 1
$ws2.Cells[2, 2].Value = "2008AE10"
$ws2.Cells[2, 3].Value = "Alice"

$ws2.Cells[3, 1].Value = 2
$ws2.Cells[3, 2].Value = "2008AE12"
$ws2.Cells[3, 3].Value = "Bob (Multiple PCNOs 5002 and 5010 -> Gets 5010)"

$ws2.Cells[4, 1].Value = 3
$ws2.Cells[4, 2].Value = "2015EF48"
$ws2.Cells[4, 3].Value = "Charlie (Multiple PCNOs 5015 and 5025 -> Gets 5025)"

$ws2.Cells[5, 1].Value = 4
$ws2.Cells[5, 2].Value = "9999ZZ99"
$ws2.Cells[5, 3].Value = "Non-existent PIS"

$pkg2.SaveAs([System.IO.FileInfo]::new("e:\Excel\test_pis_max.xlsx"))
$pkg2.Dispose()

# 3. Create test_custom_col.xlsx (testing manual mapping)
$pkg3 = New-Object OfficeOpenXml.ExcelPackage
$ws3 = $pkg3.Workbook.Worksheets.Add("Sheet1")
$ws3.Cells[1, 1].Value = "SNo"
$ws3.Cells[1, 2].Value = "EmpCode"
$ws3.Cells[1, 3].Value = "Dept"

$ws3.Cells[2, 1].Value = 1
$ws3.Cells[2, 2].Value = "5001"
$ws3.Cells[2, 3].Value = "IT"

$ws3.Cells[3, 1].Value = 2
$ws3.Cells[3, 2].Value = "5010"
$ws3.Cells[3, 3].Value = "Finance"

$pkg3.SaveAs([System.IO.FileInfo]::new("e:\Excel\test_custom_col.xlsx"))
$pkg3.Dispose()

Write-Output "All 3 test Excel files created successfully."
