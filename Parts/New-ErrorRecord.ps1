<# $myError = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList @(
    { New-Object -TypeName System.Exception },
    'ERRAppObjNotFound',
    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
    $null
) #>

$myError = [System.Management.Automation.ErrorRecord]::new([System.Management.Automation.ItemNotFoundException], 'App object not found',
  [System.Management.Automation.ErrorCategory]::ObjectNotFound, $null)
Write-Output $myError