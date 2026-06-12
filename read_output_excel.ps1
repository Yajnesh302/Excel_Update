[Reflection.Assembly]::LoadFrom("e:\Excel\bin\EPPlus.dll") | Out-Null
$pkg = New-Object OfficeOpenXml.ExcelPackage([System.IO.FileInfo]::new("e:\Excel\test_output.xlsx"))
$ws = $pkg.Workbook.Worksheets[1]
$dimension = $ws.Dimension

Write-Output "Excel Sheet Dimension: $($dimension.Address)"
Write-Output "--------------------------------------------"

for ($r = 1; $r -le $dimension.End.Row; $r++) {
    $rowText = ""
    for ($c = 1; $c -le $dimension.End.Column; $c++) {
        $rowText += "[" + $ws.Cells[$r, $c].Text + "]`t"
    }
    Write-Output $rowText
}

$pkg.Dispose()
