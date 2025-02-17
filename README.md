# GetPlacesTree.ps1

## Overview

`GetPlacesTree.ps1` is a PowerShell script that generates an XML file representing the hierarchy of Microsoft Places in a graphical format. The script connects to Microsoft Places, retrieves the list of places, and constructs a hierarchical representation in XML format. Nodes are positioned in a circle with coordinates calculated based on depth and current angle.

## Features

- Connects to Microsoft Places to retrieve the list of places.
- Constructs a hierarchical representation of places in XML format.
- Positions nodes in a circle with coordinates calculated based on depth and current angle.
- Adds nodes and edges to the XML to represent the hierarchy graphically.

## Requirements

- PowerShell 7.6.0-preview.2 or higher
- MicrosoftPlaces module version 1.2.0 or higher

## Usage

1. Ensure you have the necessary permissions to connect to Microsoft Places and have the Microsoft Places module installed.
2. Clone or download the repository.
3. Open a PowerShell terminal.
4. Navigate to the directory containing the `GetPlacesTree.ps1` script.
5. Run the script using the following command:

    ```powershell
    .\GetPlacesTree.ps1
    ```

6. Open the generated XML file (`MSPlace_hierarchy.xml`) in [draw.io](https://app.diagrams.net/) to visualize the hierarchy graphically.

## Parameters

- `placeID`: The ID of the current place.
- `placeslist`: The list of all places retrieved from Microsoft Places.
- `currentdata`: The current XML content.
- `currentDepth`: The current depth in the hierarchy.
- `currentAngle`: The current angle for node positioning.

## Example

```powershell
.\GetPlacesTree.ps1
