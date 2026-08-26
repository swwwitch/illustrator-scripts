# GenerateGuidesGrid

[![Direct](https://img.shields.io/badge/Direct%20Link-GenerateGuidesGrid.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/GenerateGuidesGrid.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenerateGuidesGrid.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script for Illustrator that divides an artboard, or the bounding box of the selection, into the specified rows and columns and generates grid guides.
- It can also draw the artboard edges, draw the cells as rectangles (with round corners and center points), and import/export presets.

<img alt="" src="https://www.dtp-transit.jp/images/ss-738-1182-72-20250713-082813.png" width="80%" />

### Main Features

- Choose the target: the selected objects, the active artboard, or all artboards
- Set rows and columns with row/column gutters (gutters can be linked)
- Set top/bottom/left/right margins individually, or link them to apply one value to all
- Extend the guides beyond the artboard by a given distance
- Draw guides on the four edges of the artboard ("Artboard Edges"); an edge with a zero margin is skipped because it would coincide with the outer grid guide
- Add a vertical guide at the horizontal center of every cell ("Split Each Cell Horizontally")
- Draw the cells as rectangles (own layer, opacity, round corners, center points)
- Clear existing guides (applied as soon as the box is checked)
- Load presets and export the current settings as a preset
- Live preview that follows every change
- Toggle between Outline and Preview view
- Automatic Japanese / English UI switching

### How to Use

1. Select the objects to target (or make the artboard active) and run the script
2. Configure rows, columns, margins, and so on in the dialog (changes appear in the live preview)
3. Click OK to commit

### Options

- **Target**: selected objects / artboard / all artboards ("Selected Object(s)" is dimmed when nothing is selected, "All Artboards" when the document has a single artboard)
- **Original object** (when the target is the selection): remove / keep / convert to guides
- **Rectangles**: draw each cell as a rectangle; opacity via the slider
- **Round corners**: apply the Round Corners live effect to each rectangle; the radius is set by value
- **Show center point**: show the center point of each rectangle in the Attributes panel (applied on OK)
- **Draw guides**: draw the grid guides
- **Guide extension**: how far the guides run past the artboard
- **Artboard edges**: draw guides on the four edges of the artboard (dimmed when the target is the selection)
- **Split each cell horizontally**: add a vertical guide at each cell's horizontal center
- **Clear existing guides**: remove the existing guides in the `grid_guides` layer (dimmed when "Draw Guides" is off)

### Notes

- Guides are created in the `grid_guides` layer and cell rectangles in the `cell-rectangle` layer; the guide layer is locked after OK.
- "Clear existing guides" only affects guides in the `grid_guides` layer. All of them are removed when the target is all artboards; otherwise only the guides on the active artboard are removed.
- A target is skipped when the margins or gutters are so large that the cell width or height would be zero or less.
- A blank or invalid row/column count is treated as 1. When the settings cannot draw a grid, existing guides are left alone even if "Clear existing guides" is checked.
- A text range selected in text-editing mode is not treated as a selected object.

### Difference from Column Setup

- The target is the artboard (Column Setup targets a path)
- The guide length is adjustable (Column Setup draws long guides)
- Top, bottom, left, and right margins can be set individually
- Settings can be stored as presets
- The generated rectangles go on their own layer at 15% opacity

### note

https://note.com/dtp_tranist/n/n7adc7290b607

### Original / Acknowledgements

Sugasawa-kun β  
https://note.com/sgswkn/n/nee8c3ec1a14c

### Update History

- v1.0.0 (20250424): Initial version
- v1.0.1 (20250427): Added guides, bleed guides, and preset export feature
- v1.7.1 (20260827): Added "Artboard Edges"; clearing existing guides now applies as soon as it is checked and is scoped to the target; fixed duplicate guides, empty-field errors, and modal preview alerts; fixed preset unit conversion; reorganized the UI layout and naming
- v1.7.2 (20260827): Artboard edges are no longer drawn on a zero-margin side; a blank row/column count is treated as 1; existing guides are kept when nothing can be drawn; "Clear Existing Guides" is disabled while "Draw Guides" is off; fixed the enabled state of the fields after loading a preset; widened the numeric fields so converted units fit; fixed the notification and rollback when a preview step fails
