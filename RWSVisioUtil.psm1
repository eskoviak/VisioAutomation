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

Add-Type -TypeDefinition @"
  public enum VisioError {
    ERRNoObjFound
  }
"@

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

##############################
#.SYNOPSIS
# Gets the visio.application object if one is loaded, returns friendly COMException if not found
#
#.DESCRIPTION
# Searches the object table for a running instance of visio, [Microsoft.Office.Interop.Visio.Application].  This is a private 
# function meant to be called internally only.  If an object is found, a handle to it is returned.  Note that if no documents are
# loaded, MOI.Visio.Applicaton.Documents object will be empty.
#
#.EXAMPLE
#An example
#
#.NOTES
#General notes
##############################
function Get-VisioObject () {
  try {
    $visio = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
    # Since no COMException was thrown a visio application object is found in the object store
    Write-Output $visio
  } catch [System.Runtime.InteropServices.COMException] {
    Write-Host('No Excel Objects Found')
    throw 'ERRNoObjFound'
  }
}

#### End PRIVATE Functions