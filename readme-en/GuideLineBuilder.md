# Build construction lines for logo design and analysis

[![Direct](https://img.shields.io/badge/Direct%20Link-GuideLineBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/GuideLineBuilder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GuideLineBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Collects the straight segments of the selection and extends each one across the drawing area. It is meant for designing and analysing logos and icons, where it makes the lines behind the shapes visible.

Groups and compound paths are walked recursively. Text is handled by outlining a temporary copy, so you can run the script with the text still selected — the original text is never changed.

### Features

**Construction lines**

- Straight segments: draws the extension of every straight segment. Turn it off to draw nothing but the circles from "Create circles from arcs"
- Horizontal / Vertical / Diagonal lines: narrows down which directions are drawn. Option (Alt) + click leaves only the direction you clicked on
- Create circles from arcs: estimates a circle from a Bézier segment. When the segment is not a true arc, the fallback is Ignore, Chord, or Extend chord
- Stroke Width: follows the "strokeUnits" preference and starts at the equivalent of 0.1 mm. Arrow keys step by 1, shift by 10, and option by 0.1

**Shapes on anchor points**

- Places a Circle or a Square at every anchor point
- Size follows the ruler unit ("rulerType") and starts at the equivalent of 1 mm
- The color is either Black (K100) or Blue (R78 G128 B255)
- Anchor points that sit on the same position get a single shape

**Preview and zoom**

- Preview is on by default: everything is drawn on a temporary layer and removed as soon as the dialog closes
- The slider at the bottom of the dialog changes the zoom level. Option centers on the artboard, shift centers on the selection, and option + shift fits the selection in the window
- With Light Preview on, the view is only updated when you release the slider

### Usage

1. Select the objects.
2. Run the script.
3. Set the options and click OK.

### Options

- Group output: puts the construction lines and the anchor shapes into groups of their own
- Use separate layer: sends the lines to a `_construction_guide` layer and the anchor shapes to a `_construction_anchorpoint` layer
- Convert to guides: creates the construction lines as guides (anchor shapes never become guides)
- Remove duplicate lines: draws only one line for any given pair of endpoints

### Notes

- When the selection does not intersect the active artboard, a virtual rectangle four times the selection's width and height is centered on it, and the lines are drawn within that.
- Running the script again with "Use separate layer" on removes only the items this script generated. Anything you added by hand on those layers is kept.
- If the selection is already on the `_construction_guide` layer, the existing one is renamed to `_construction_guide_backup...` and a fresh layer is created; the original objects are kept.

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/nd801b9b0367f)

### Update History

- v1.2.1 (2026-09-17): Preview is now on by default. Fixed the dialog position not being remembered, the output layer changing after using the preview, and the previous anchor shapes surviving when the shape is set back to None. Code cleanup as well
- v1.2 (2026-03-12)
