$users = Import-Csv "Userslistfile.txt" -Header 'DC','User','WMS','Area'

$excel = New-Object -ComObject Excel.application
$excel.visible = $true
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open("Alphapasswords.xlsx")
$worksheet = $workbook.worksheets(1)
$worksheet.Activate()
$lastrow = $worksheet.UsedRange.rows.count + 1



foreach ($data in $users){

$name = $data.User

$worksheet.Cells.item($lastrow,1) = $data.DC

if ($data.WMS.count -lt 10){
$worksheet.cells.item($lastrow,2) = "WMS00" + $data.WMS
}

else { $worksheet.cells.item($lastrow,2) = $data.WMS}

$worksheet.cells.item($lastrow,3) = $data.User

$worksheet.Cells.item($lastrow,4) = Read-Host Choose a password for $name

$worksheet.cells.item($lastrow,5) = $data.Area

$worksheet.cells.item($lastrow,6) = Read-Host Heat number for $name

$lastrow = $lastrow + 1

}

$workbook.SaveAs("Alphapasswords.xlsx")
$excel.Quit()
