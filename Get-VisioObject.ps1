
try {
    $visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
    Write-Output('Found {0}' -f $visio)    
} catch [System.Runtime.InteropServices.COMException] {
    Write-Output('No Excel Objects Found')
    return -1
}
