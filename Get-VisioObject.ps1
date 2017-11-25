
try {
    $visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
    # Since no COMException was thrown a visio application object is found in the object store
    # If the Documents collection is empty, there are no documents loaded.  Otherwise, it has objects
    # representing the .vsdx and .vssx files that have been loaded
    if($visio.Documents.Count -ne 0) {
        foreach ($doc in $visio.Documents) {
            if ($doc.Name.IndexOf('.vsdx') -ge 0) {
                Write-Output($doc.Name)
            }
        }
    } else {
        Write-Host('No Documents loaded')
    }
} catch [System.Runtime.InteropServices.COMException] {
    Write-Host('No Excel Objects Found')
    return -1
}
