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

- Pick [None] or a favorite arrowhead (Arrow 8, 11, 27) with icons showing their shapes (the selected one has a gray background). [None] at the top (a plain line) removes the arrowheads from both ends
-  Arrow 11 is selected at start. The scale changes to suit it (100% for 11 and 27, 25% for 8). The tip alignment changes too (at end of path for 8 and 11, beyond end of path for 27)
- Pick any other arrowhead from the pop-up menu (scale: 100%)
- Options (icons in one row; names are in the tooltips; dimmed while the arrowhead is [None])
  - **Same at end** (link icon): puts the same arrowhead on the end (the arrowhead icons show both ends too). When off, the end has no arrowhead
  - **Swap start and end** (⇄ icon): puts the arrowhead on the end instead of the start (dimmed while Same at end is on)
  - Tip alignment (left: beyond end of path / right: at end of path)

#### Dashes

- Choose **None**, **Dashed** or **Dotted** with radio buttons. Choosing Dashed or Dotted fills in Segments, Gap and Dash based on the stroke weight, and sets the arrowhead to [None]
- Dash Calculation: works out the dashes from Segments, Gap and Dash. With several paths selected, each path is calculated from its own length
- Calculation: Gap→Dash / Dash→Gap
- Dotted works out the gap between zero-length dots from Segments and fixes the cap to Round
- Adjust ends (picked with icons; left: keep dash lengths / right: adjust ends): on an open path, Adjust ends distributes the dashes so both ends finish with a dash (or dot)

#### Preview

- **Open Stroke Panel** at the bottom left closes the dialog (reverting like Cancel) and opens Illustrator's Stroke panel
- The preview is always on (from the moment the dialog opens). The weight updates live while you type; settings that include arrowheads are previewed by playing the action and undoing it. Cancel restores the original state

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
- v1.1.1 (2026-10-04) Same at end and Swap start and end moved into Options in the Arrowheads panel. Gap and Dash now show their unit (pt) inside the field. Arrow 1 moved from the favorites to the pop-up menu. Picking a favorite arrowhead also sets the tip alignment. Calculation and Adjust ends moved into the Dash Calculation panel
- v1.1.2 (2026-10-04) Options are now a row of icons instead of a panel (link icon for Same at end, ⇄ for Swap start and end, two icons for tip alignment). Tightened the gap between the favorite arrowheads and the pop-up menu
- v1.1.3 (2026-10-04) The icons added for tip alignment in v1.1.2 now belong to Adjust ends (keep dash lengths / adjust ends). Tip alignment got its own icons (beyond end of path / at end of path). Fixed the last dot sometimes missing on dotted lines with Adjust ends. Options are dimmed while the arrowhead is [None]. Choosing Dashed or Dotted sets the arrowhead to [None]. Larger option icons, and the Adjust ends icons now have space between the frame and the shapes. [None] and the favorite arrowheads are now icons instead of radio buttons. The preview is always on; the Preview checkbox is replaced by an Open Stroke Panel button. The scale shows its % inside the field

### Script info

- Version: v1.1.3
- First release: 2026-10-03
- Last updated: 2026-10-04
