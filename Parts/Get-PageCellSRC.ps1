$visio = .\Get-VisioObject.ps1

$page = $visio.Documents['VisioDOM.vsdx'].Pages['MOIV.Page']

$shape = $page.Shapes['MOI.Visio.Shape.Master']

$pinX = $shape.CellsSRC([Microsoft.Office.Interop.Visio.VisSectionIndices]::visSectionObject,
  [Microsoft.Office.Interop.Visio.VisRowIndices]::visRowXFormOut,
  [Microsoft.Office.Interop.Visio.VisCellIndices]::visXFormPinX)



