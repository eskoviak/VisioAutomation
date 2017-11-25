<#
  Module containing Visio automations for the EA Diagram Process
#>

#### Begin PUBLIC Funtions
function New-Diagram {
    param(
    # Bring in the PowerShell defaults
    [CmdletBinding()]
    [Parameter()]
    [String]$DiagramName
    ) 
    
}

function Get-VisioObject {
  param (
    # Bring in PS Defaults
    [CmdletBinding()]
    [Parameter()]
    [String]$DocumentName = ""
    
  )
  
}

Export-ModuleMember
#### End PUBLIC Funtions

#### Begin PRIVATE Funtions

#### End PRIVATE Functions