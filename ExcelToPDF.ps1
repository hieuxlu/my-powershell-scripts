$path = Get-Location
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]
# Add -recurse to get files in all sub-folders
$excelFiles = Get-ChildItem -Path ($path.Path + "\*") -include *.xls, *.xlsx 

$objExcel = New-Object -ComObject excel.application
$objExcel.visible = $false
foreach ($wb in $excelFiles) {
    $filepath = Join-Path -Path $path -ChildPath ($wb.BaseName + ".pdf")
    $workbook = $objExcel.workbooks.open($wb.fullname, 3)
    $workbook.Saved = $true
    "saving $filepath"
    $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)
    $objExcel.Workbooks.close()
}
$objExcel.Quit()