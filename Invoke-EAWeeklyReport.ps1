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
$range = $EAWorkSheet.UsedRange
#Write-Host ($range.Address())
foreach($row in $range.Rows) {
    write-host $("Row: {0}; ???: {1}" -f $row.Row, $range.Item($row.Row,1).Value())
}

# Close down Excel Gracefully and Completely
$EAWorkbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

