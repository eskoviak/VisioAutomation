#Import-Module ImportExcel
[cmdletbinding()]
param()

#If testing internally, use the Test file on the EA site; If testing externally, use the copy on the external SharePoint
$computerName = Get-WmiObject Win32_ComputerSystem | Select-Object Name
if ($computerName.Name -like "EDSKOV-HP840") {$EAProjectListWorkbook = "http://rede.rwsc.net/teams/IT/EA/Documents/EA%20Project%20List%20Test.xlsx"}
elseif ($computerName.Name -like "EdGamer") {$EAProjectListWorkbook = "http://redwingshoes.sharepoint.com/Shared%20Documents/EA%20Project%20List%20Test.xlsx"}
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

#lamba function to get the current cell
$currentCell = { param($row, $column) $EAWorkSheet.Cells($row, $column)}

# get the column headers
$column = 1
$headers = New-Object -TypeName System.Collections.ArrayList
$projectList = @()
while ($currentCell.Invoke(1, $column).Text() -ne [string]::Empty ) {
    $temp = $($currentCell.Invoke(1, $column).Text() -replace "\s+", " ")
    $temp = $temp -replace "^\s", ""
    $headers.Add($temp)
    $column += 1
}

# Create the row template
$EAProjectItemRowTempate = New-Object psobject
foreach($colHead in $headers) {
    Add-Member -InputObject $EAProjectItemRowTempate -MemberType NoteProperty -Name $colHead -Value ""
}
Write-Verbose $EAProjectItemRowTempate

$firstRow = $true
foreach($row in $EAWorkSheet.UsedRange.Rows) {
    if($firstRow) {
        $firstRow = $false
        continue
    }

    $column = 1
    foreach($colHead in $headers){
        $EAProjectItemRowTempate.($colHead) = $row.Cells(1, $column).Text()
        $column += 1
    }

    $projectList += $EAProjectItemRowTempate.psobject.Copy()
}

Write-Host $projectList | select 'BT Team Primary'

# Close down Excel Gracefully and Completely
$EAWorkbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

