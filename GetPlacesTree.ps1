<#
.SYNOPSIS
    This script generates an XML file that represents the hierarchy of Microsoft Places in a graphical format.

.DESCRIPTION
    The script connects to Microsoft Places, retrieves the list of places, and constructs a hierarchical representation in XML format.
    It uses functions to add nodes and edges to the XML and to recursively explore child nodes.
    Nodes are positioned in a circle with coordinates calculated based on depth and current angle.

.PARAMETER placeID
    The ID of the current place.

.PARAMETER placeslist
    The list of all places retrieved from Microsoft Places.

.PARAMETER currentdata
    The current XML content.

.PARAMETER currentDepth
    The current depth in the hierarchy.

.PARAMETER currentAngle
    The current angle for node positioning.

.EXAMPLE
    .\GetPlacesTree.ps1
    Runs the script and generates the XML file with the hierarchy of places.

.NOTES
    Ensure you have the necessary permissions to connect to Microsoft Places and have the Microsoft Places module installed.

.REQUIREMENTS
    - PowerShell 7.6.0-preview.2 or higher
    - MicrosoftPlaces module version 1.2.0 or higher
#>

# Check PowerShell version
if ($PSVersionTable.PSVersion -lt [version]"7.6.0-preview.2") {
    exit
}

# Check and install necessary modules
$modules = @("MicrosoftPlaces")
$minVersion = "1.2.0"

foreach ($module in $modules) {
    if ((Get-Module -ListAvailable -Name $module).Version -lt $minVersion) {
        Install-Module -Name $module -RequiredVersion $minVersion -Force -AllowClobber -Scope CurrentUser -Verbose
    }
}

# Import necessary modules
Import-Module MicrosoftPlaces

# Creation of service functions

# adding a node to the XML

function Add-Node {
    param (
        [string]$placeID,
        [string]$placeDisplayname,
        [string]$placeParentID,
        [string]$placeX,
        [string]$placeY
    )

    if($placeParentID -eq ""){
        $placeParentID = 1
    }

    $node = @" 
                <mxCell id="$placeID" value="$placeDisplayname" style="ellipse;whiteSpace=wrap;html=1;rounded=1;shadow=1;comic=0;labelBackgroundColor=none;strokeWidth=2;fontFamily=Verdana;fontSize=12;align=center;" parent="$placeParentID" vertex="1">
                    <mxGeometry x="$placeX" y="$placeY" width="90" height="90" as="geometry" />
                </mxCell>
"@
    return $node
}

# adding an edge to the XML

function Add-Edge {
    param (
        [string]$placeID,
        [string]$placeParentID
    )

    $random = Get-random
    $arrowID ="$($placeID)-$($random)"
    $edge = @"
                <mxCell id="$arrowID" style="edgeStyle=none;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;" edge="1" parent="1" source="$placeParentID" target="$placeID">
                    <mxGeometry relative="1" as="geometry" />
                </mxCell>
"@
    return $edge
}

# searching for child nodes

function Search-ChildNode {
    param (
        [string]$placeID,
        $placeslist,
        [string]$currentdata,
        [int]$currentDepth,
        [int]$currentAngle
    )

    # increment depth
    $currentDepth++
    $currentAngle += 180
    # retrieve list of child nodes
    $childnode = $placeslist | where-object {$_.ParentID -eq $placeID} | Sort-Object type, displayname
    
    if($childnode){
        
        # if child nodes exist
       
        # create graphic variables
        if($placeID -eq ""){
            $nVertex = $childnode.count
        } else {
            $nVertex = $childnode.count + 1
        }
        
        $angle = 360 / $nVertex

        # explore child nodes
        foreach($child in $childnode){
            $childID = $child.PlaceId
            $childDisplayname = $child.DisplayName.Replace("&","-")
            $currentAngle += $angle
            $result = Search-ChildNode -placeID $childID -placeslist $placeslist -currentdata $currentdata -currentDepth $currentDepth -currentAngle $currentAngle
            $reachedDepth = $result[0]
            $currentdata = $result[1]
            
            $exp = $reachedDepth - $currentDepth
            $r = 50*[MATH]::Pow(2,$exp)
            $angleRadians = $currentAngle * [math]::PI / 180
            $sin = [math]::Sin($angleRadians)
            $cos = [math]::Cos($angleRadians)
            $x = [math]::Round($r * $cos)
            $y = [math]::Round($r * $sin)
            
            $currentdata += Add-Node -placeID $childID -placeDisplayname $childDisplayname -placeParentID $placeID -placeX $x -placeY $y
            $currentdata += Add-Edge -placeID $childID -placeParentID $placeID

        }
    } else {
        
        # if no child nodes exist
        $reachedDepth = $currentDepth
        
    }

    return $reachedDepth,$currentdata
}

# Main

Connect-MicrosoftPlaces

$places = get-placev3 | where-object {$_.Type -ne "roomlist"} | Sort-Object type, displayname

# Output file name
$xmloutputfile = ".\MSPlace_hierarchy.xml"
# Static header for text file
$cdataContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="Electron" agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) draw.io/26.0.9 Chrome/128.0.6613.186 Electron/32.2.5 Safari/537.36" version="26.0.9">
    <diagram name="Page-1" id="ad52d381-51e7-2e0d-a935-2d0ddd2fd229">
        <mxGraphModel dx="$dx" dy="$dy" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" background="none" math="0" shadow="0">
            <root>
                <mxCell id="0"/>
                <mxCell id="1" parent="0"/>
"@

# Add nodes from previous tables

$resultchild = Search-ChildNode -placeID "" -placeslist $places -currentdata $cdataContent -currentDepth -1 -currentAngle 180

$x = $resultchild[0]
$cdataContent = $resultchild[1]

# Close the XML model
$cdatacontent += @"
            </root>
        </mxGraphModel>
    </diagram>
</mxfile>
"@

$cdataContent > $xmloutputfile

if(Test-Path($xmloutputfile)){
    "JOB COMPLETE!"
} else {
    "JOB FAILED!"
}