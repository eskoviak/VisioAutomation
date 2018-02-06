#Import-Module ImportExcel

#If testing internally, use the Test file on the EA site; If testing externally, use the copy on the external SharePoint
$computerName = Get-WmiObject Win32_ComputerSystem | Select-Object Name
if ($computerName.Name -like "EDSKOV-HP840") {$EAProjectListWorkbook = "http://rede.rwsc.net/teams/IT/EA/Documents/EA%20Project%20List%20Test.xlsx"}
elseif ($computerName.Name -like "EdGamer") {$EAProjectListWorkbook = "http://redwingshoes.sharepoint.com/Shared%20Documents/EA%20Project%20List.xlsx"}
else { 
    Write-Error ("Unknown computer Name {0}" -f $computerName.name)b
    Exit
}

# the name of the workbook
$EAProjectListStatus = "Status"  # The name of the sheet with the status

# Do It!
$excel = New-Object -ComObject Excel.Application
$EAWorkbook = $excel.Workbooks.Open($EAProjectListWorkbook)
$EAWorkSheet = $EAWorkbook.Sheets.Item($EAProjectListStatus)
# Keep the application hidden
$excel.Visible = $false
#$range = $EAWorkSheet.UsedRange

# Get the column Headers
$column = 1
$currentCell = { param($row, $column) $EAWorkSheet.Cells($row, $column)}

while ($currentCell.Invoke(1, $column).Text() -ne [string]::Empty ) {
    Write-Host ("Heading: {0}" -f $currentCell.Invoke(1, $column).Text())
    $column += 1
}

#Write-Host ("A1: {0}" -f $EAWorkSheet.Cells(1,1).Value())


#Write-Host ("Range Address {0}, Range End {1}" -f $range.Address(), $range.End([Microsoft.Office.Interop.Excel.XlDirection]::xlToRight).Address())
<# foreach($row in $range.Rows) {
    write-host $("Row: {0}; ???: {1}" -f $row.Row, $range.Item($row.Row,1).Value())
} #>

# Close down Excel Gracefully and Completely
$EAWorkbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

