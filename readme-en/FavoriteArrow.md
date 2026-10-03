# Apply a favorite arrowhead and stroke settings at once

[![Direct](https://img.shields.io/badge/Direct%20Link-FavoriteArrow.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/FavoriteArrow.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteArrow.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Applies a favorite arrowhead together with the stroke settings — weight, cap, corner and dashes — to the selected paths.
- Arrowheads cannot be reached from the Illustrator DOM, so a temporary action (ai_plugin_setStroke) is generated and played instead.
- An adaptation of [SetStrokeAndArrowheads](SetStrokeAndArrowheads.md) that adds the cap and corner settings from SetStrokeAlignment and the dash calculation from [DashGapCalculator](DashGapCalculator.md).

### Features

#### Presets

- The pop-up at the top of the dialog loads saved settings (stroke, arrowheads and dashes)
- **Save...** names and saves the current settings (an existing name is overwritten after confirming); **Delete** removes the selected preset
- Presets survive an Illustrator restart (Folder.userData/illustrator-scripts/FavoriteArrowPresets.json)

#### Stroke

- Weight
- Cap (Butt / Round / Projecting) and corner (Miter / Round / Bevel)

#### Arrowheads

- [None] at the top removes the arrowheads from both ends
- Pick a favorite arrowhead (Arrow 1, 8, 11, 27) with a radio button. Arrow 11 is selected at start. The scale changes to suit it (100% for 1, 11 and 27, 25% for 8)
- Pick any other arrowhead from the pop-up menu (scale: 100%)
- **Same at end**: puts the same arrowhead on the end. When off, the end has no arrowhead
- **Swap start and end**: puts the arrowhead on the end instead of the start (dimmed while Same at end is on)
- **Options**: tip alignment (at end of path / beyond end of path)

#### Dashes

- Choose **None**, **Dashed** or **Dotted** with radio buttons. Choosing Dashed or Dotted fills in Segments, Gap and Dash based on the stroke weight
- Dash Calculation: works out the dashes from Segments, Gap and Dash. With several paths selected, each path is calculated from its own length
- Calculation: Gap→Dash / Dash→Gap
- Dotted works out the gap between zero-length dots from Segments and fixes the cap to Round
- **Adjust ends**: on an open path, distributes the dashes so both ends finish with a dash (or dot)

#### Preview

- The weight updates live while you type; settings that include arrowheads are previewed by playing the action and undoing it. Cancel restores the original state

### Usage

1. Select the paths to style (paths inside groups and compound paths are included)
2. Run the script and set the stroke, arrowheads and dashes in the dialog (or pick a saved preset)
3. Click **OK** to apply

### Notes

- Arrowhead, tip alignment, cap and corner names must match Illustrator's UI labels (they depend on the UI language).
- The arrowhead scale keys (asc1 / asc2) are estimated.
- Dashes are set through the DOM after the action runs.
- The favorite arrowheads and their scales can be changed in `FAVORITE_ARROWS` at the top of the script.

---

### Update History

- v1.0.0 (2026-10-03) Initial release
- v1.0.1 (2026-10-04) The initial stroke weight now follows the general unit (0.25 pt for mm, 1 px for px, 5 pt otherwise)
- v1.0.2 (2026-10-04) The cap and corner now start at Butt and Miter. Arrow 27 joins the favorite arrowheads
- v1.1.0 (2026-10-04) Added presets (save, load, delete). Added [None] at the top of the favorite arrowheads to remove arrowheads. Dashed and Dotted are now None / Dashed / Dotted radio buttons instead of checkboxes. Tip alignment moved into Options in the Arrowheads panel. Arrow 11 is now selected at start

### Script info

- Version: v1.1.0
- First release: 2026-10-03
- Last updated: 2026-10-04
