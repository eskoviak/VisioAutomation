
try {
    $visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
    # Since no COMException was thrown a visio application object is found in the object store
    Write-Output $visio
} catch [System.Runtime.InteropServices.COMException] {
    Write-Host('No Excel Objects Found')
    throw 'ERRNoObjFound'
}
