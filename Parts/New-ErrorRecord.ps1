<# $myError = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList @(
    { New-Object -TypeName System.Exception },
    'ERRAppObjNotFound',
    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
    $null
) #>

$myError = [System.Management.Automation.ErrorRecord]::new([ObjectNotFoundException], 'App object not found', 2, [ErrorCategory]::ObjectNotFound, $null)
Write-Output $myError