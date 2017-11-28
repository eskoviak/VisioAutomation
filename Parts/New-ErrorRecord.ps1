<# $myError = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList @(
    { New-Object -TypeName System.Exception },
    'ERRAppObjNotFound',
    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
    $null
) #>

$myError = [System.Management.Automation.ErrorRecord]::new(
  [System.Exception]::new('String'),
  'App object not found',
  [System.Management.Automation.ErrorCategory]::ObjectNotFound,
  {}
  )
Write-Output $myError