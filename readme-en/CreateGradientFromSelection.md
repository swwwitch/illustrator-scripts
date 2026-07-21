# CreateGradientFromSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-CreateGradientFromSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/CreateGradientFromSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Extracts fill and stroke colors from the selection in layout order (left to right for horizontal rows, top to bottom for vertical stacks) and registers them in a swatch group.
- Builds a linear gradient from the extracted colors, optionally applying it to a new rectangle or saving it as a Graphic Style.
- When nothing is selected, two or more swatches selected in the Swatches panel can be used as the input instead.

### Main Features

- Recursively walks groups, compound paths, and text, collecting both fill and stroke colors (duplicates removed)
- **Make Global colors**: registers swatches as Global Process colors so they can be edited together later
- **Create gradient**: builds a linear gradient with the extracted colors assigned to its stops
  - **Smooth**: colors evenly spaced across the ramp
  - **Segmented**: hard-edged gradient (2 to 6 colors; disabled when 7 or more items are selected)
- **Create rectangle and apply gradient**: draws a rectangle filled with the new gradient
  - **Match selection size**: uses the bounding size of the selection (otherwise 100 pt square; fixed 200 x 100 pt when invoked from swatches)
  - Placed below and left-aligned for a horizontal selection, to the right and top-aligned for a vertical selection
- **Save as Graphic Style**: registers the rectangle's appearance as a Graphic Style (when rectangle output is off, a temporary rectangle on a temporary layer is used and removed afterwards)
- When the selection is detected as vertical, an action (gradient/90degree) rotates the gradient angle to 90 degrees
- Automatic Japanese / English UI, with tooltips on the key options
- Dialog values persist for the session inside the targetengine (cleared when Illustrator restarts)

### Workflow

1. Collect colors from the selected objects (or selected swatches) and detect the layout orientation
2. Set the color and rectangle options in the dialog
3. Create a new swatch group (AutoGradient) and register the extracted colors as AutoColor swatches
4. Build the gradient, then place the rectangle and register the Graphic Style as requested

### Not Supported

- No open document (exits silently)
- Fewer than two distinct colors found (exits silently)
- Empty selection combined with fewer than two selected swatches
- Objects with neither fill nor stroke, and objects whose bounds cannot be read
- No rectangle is created when every layer is locked or hidden

### Update History

- v1.9.2 (20260528): Current version
