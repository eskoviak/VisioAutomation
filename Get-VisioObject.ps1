$visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')

Write-Output('Found {0}' -f $visio)