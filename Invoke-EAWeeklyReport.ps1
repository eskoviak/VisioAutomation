#Import-Module ImportExcel

$EAProjectListWorkbook = "http://redwingshoes.sharepoint.com/Shared%20Documents/EA%20Project%20List.xlsx"
#$EAProjectListWorkbook = "http://rede.rwsc.net/teams/IT/EA/Documents/EA%20Project%20List.xlsx" # the name of the workbook
$EAProjectListStatus = "Status"  # The name of the sheet with the status

#$EAProjectList = Import-Excel -Path $EAProjectListWorkbook -WorksheetName $EAProjectListStatus
#$EAProjectList | ForEach-Object {
#    Write-Host $_.'BT Team'
#}

$excel = New-Object -ComObject Excel.Application
$EAWorkbook = $excel.Workbooks.Open($EAProjectListWorkbook)
$EAWorkSheet = $EAWorkbook.Sheets.Item($EAProjectListStatus)
$excel.Visible = $false
#write-host $($EAWorkSheet.UsedRange)
$range = $EAWorkSheet.UsedRange
#Write-Host $range.Row
foreach($row in $range.Rows) {
    write-host $("Row: {0}; ???: {1}" -f $row.Row, $($row.Item(1).Value[1]))
}
$EAWorkbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

