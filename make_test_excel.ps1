[Reflection.Assembly]::LoadFrom("e:\Excel\bin\EPPlus.dll") | Out-Null
$pkg = New-Object OfficeOpenXml.ExcelPackage
$ws = $pkg.Workbook.Worksheets.Add("Sheet1")

# Headers
$ws.Cells[1, 1].Value = "Serial Number"
$ws.Cells[1, 2].Value = "PIS NO"
$ws.Cells[1, 3].Value = "Name"
$ws.Cells[1, 4].Value = "Rank"

# Row 2 (matches PIS 1001)
$ws.Cells[2, 1].Value = 1
$ws.Cells[2, 2].Value = "1001"
$ws.Cells[2, 3].Value = "John Doe"
$ws.Cells[2, 4].Value = "Manager"

# Row 3 (matches PIS 1002)
$ws.Cells[3, 1].Value = 2
$ws.Cells[3, 2].Value = "1002"
$ws.Cells[3, 3].Value = "Jane Smith"
$ws.Cells[3, 4].Value = "Officer"

# Row 4 (non-matching PIS 9999)
$ws.Cells[4, 1].Value = 3
$ws.Cells[4, 2].Value = "9999"
$ws.Cells[4, 3].Value = "Stranger"
$ws.Cells[4, 4].Value = "Unknown"

$pkg.SaveAs([System.IO.FileInfo]::new("e:\Excel\test_input.xlsx"))
$pkg.Dispose()
Write-Output "Excel test input file created."
