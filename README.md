# Generate Azure Stencils from SVG Icons

## Overview

This tool generates Microsoft Visio `.vssx` stencil files from folders of Azure SVG icon files.  
Each SVG becomes a Visio master shape with:

- Standardized size
- Editable label text
- Connection points around the symbol (optionally with outward offsets) for clean connector snapping

The output can be used directly in Visio under *My Shapes*.

## Features

- Automatic stencil creation per SVG folder (`Azure-<Category>.vssx`)
- Imports each SVG into a Visio master and scales width to `0.7 in` (maintaining aspect ratio)
- Detects and uses the top-level group shape when present
- Creates a label from the SVG file name with configurable font family and size
- Removes label border and fill, and positions the label below the symbol
- Adds 12 connection points (3 per side) with optional outward offsets (e.g. `0.05 in`)
- Fixes connection points (`visSLOFixedNoFoldToShape`) so they do not fold back to the shape edges
- Clears double\-click events on child shapes and enables text editing on the label (`=OPENTEXTWIN()`)

## Core Logic

The main work is done in a `VisioStencilGenerator` class, which:

1. Manages the Visio application lifetime (runs Visio hidden)
2. Groups SVG files by folder (category)
3. For each category, creates a new Visio stencil document
4. For each SVG file:
    - Creates a master
    - Imports the SVG
    - Scales the master to the standard width
    - Adds a text label under the icon
    - Adds connection points around the shape
5. Saves the resulting `.vssx` files to the configured output folder

## Connection Points

Connection points are defined using Visio shape formulas based on `Width` and `Height`, plus offsets.

Typical layout:

- Left side: 3 points at `0.25\*Height`, `0.5\*Height`, `0.75\*Height`, optionally with a small negative X offset (e.g. `-0.05 in`) to place them slightly outside the left edge
- Right side: 3 points at `0.25\*Height`, `0.5\*Height`, `0.75\*Height`, optionally at `Width+0.05 in`
- Bottom side: 3 points at `0.25\*Width`, `0.5\*Width`, `0.75\*Width`, optionally at `-0.05 in`
- Top side: 3 points at `0.25\*Width`, `0.5\*Width`, `0.75\*Width`, optionally at `Height+0.05 in`

Examples of formulas:

- Left, 25% up: `-0.05 in`, `0.25\*Height`
- Top, 25% across: `0.25\*Width`, `Height+0.05 in`
- A point above the top edge: `0.25\*Width`, `Height+0.1 in`
- A point outside the left edge: `-0.1 in`, `0.5\*Height`

Visio supports negative coordinates and positive offsets beyond `Width`/`Height`, so these can be used to move connectors slightly away from the symbol.

## Label Formatting

Label shapes (text under each icon) are formatted to be consistent:

- Text height is fixed (e.g. `0.25 in`)
- Horizontally centered under the symbol
- Positioned below the main shape (e.g. `TxtPinY = -0.30`)
- No border or fill on the label shape
- Double\-click opens the text editor (`=OPENTEXTWIN()`), so labels are easy to rename

## Configuration via `appsettings.json`

The application is configured using a JSON file named `appsettings.json` in the project directory.

### Structure

A typical `appsettings.json` might look like:

```json
{
  "Paths": {
    "IconsRoot": "C:\Temp\AzureIcons",
    "OutputFolder": "%USERPROFILE%\Documents\My Shapes"
  },
  "Font": {
    "Type": "Segoe UI",
    "Size": 10
  }
}

### How to run

1. Ensure `appsettings.json` is present in the project root and configured with valid values for:
   - `Paths:IconsRoot`
   - `Paths:OutputFolder`
   - `Font:Type`
   - `Font:Size`

2. Build the project from the project root:

   ```bash
   dotnet build

3. Run the application:

   ```bash
   dotnet run --project AzureVisioStencilGenerator.csproj
