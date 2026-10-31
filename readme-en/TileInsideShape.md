# Tile the smaller object inside the shape of the larger one

[![Direct](https://img.shields.io/badge/Direct%20Link-TileInsideShape.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TileInsideShape.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TileInsideShape.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

With two or more objects selected, duplicates the smaller ones inside the shape of the largest one.
Copies that stick out of the shape are removed (or kept under a mask), and the originals are left in place.

<img alt="The Tile Inside Shape dialog" src="../png/ss-724-1090-144-20261027-130407-s.png" width="35%" />

### Features

- Only copies that fit inside the path are kept. Works with curves, concave shapes and holes in compound paths. Groups use the shape of each object in the group (a clipping group uses its clipping path). Text uses the shapes of its outlined glyphs (a duplicate is outlined for measuring and then deleted, so the text itself is unchanged). Other objects such as placed images use their bounding box
- With Keep Overflow and Mask, tiles that stick out of the shape are kept and masked with the fill area's shape (masks for groups and text are compound paths of flattened outlines)
- Choose the layout from Grid, Brick (Horizontal), Brick (Vertical), Hexagonal, Random (In Order) and Random (Random Order) (Brick (Horizontal) by default). Hexagonal closes up the rows to pack round tiles tightly. The two Random layouts place tiles without overlapping, keeping the spacing, and Density sets how many to keep. Reshuffle draws a new arrangement, jitter and order
- Center lines the arrangement up with the center of the fill area so the margins are even (on by default; a tile or a gap goes at the center, whichever fits more)
- Angle rotates the whole arrangement together with the tiles
- Spacing X and Spacing Y can be set separately (linked with the link icon)
- Adjust X and Adjust Y shift the tile positions (positive values move right and down; the slider's center leaves them in place and each end shifts by half the pattern period)
- The Rotate Each panel sets Rotate and Flip Every Other 180°. Rotation and Size in the number settings add random jitter to each tile's rotation and size (0 means none)
- Text can be tiled too. Its size, spacing and fit use the glyphs rather than the text box, and the glyphs are centered on each spot (also for text inside groups)
- With several objects to tile selected, they are used in turn from the frontmost. Only Random (Random Order) picks them at random
- Scale, Spacing, Offset, Adjust X, Adjust Y, Angle, the Rotation and Size jitter and the random Density can be changed with either the field or the slider. Length units follow the current ruler settings
- Fit inside the fill area's path offset by a set amount (negative shrinks inward, positive lets tiles stick out)
- The number to the right of the Scale slider shows how many tiles the current settings place. Preview can be turned off
- The original objects are left in place. With Fill Area, its opacity can be lowered, or it can be deleted or turned into guides
- Duplication as a symbol (on by default)
- Remembers the settings confirmed with OK until Illustrator quits and uses them as the defaults next time

### Usage

1. Select the fill area and one or more objects to tile. The largest object becomes the fill area.
2. Run the script.
3. Set the scale, spacing and other options, then run it.

### Article

[Introduction article (Japanese)｜DTP Transit 別館](https://note.com/dtp_tranist/n/ne393f59cbb77)

### Update History

- v2.0.1 (2026-10-31) When an object to tile is text (or a group containing text), its size, spacing and fit now use the glyphs rather than the text box, and the glyphs are centered on each spot (point text used to be measured by its advance and line height, and area text by its frame, which widened the spacing and shifted the positions). Turned the Tiles panel into the Rotate Each panel (Rotate, Flip Every Other 180°) and removed the Jitter checkbox, so the Rotation and Size jitter can always be set (0 means none); merged Order into the Layout choices Random (In Order) and Random (Random Order) (Grid, Brick and Hexagonal use the objects in turn from the frontmost); moved the Rotation and Size jitter, now with sliders, below Angle and the Tiles count, without its label, to the right of the Scale slider; laid out more compactly so the dialog is narrower (the Rotate Each checkbox renamed to Rotate, and Flip Every Other 180° no longer indented). Shortened the sliders. The Scale slider now goes up to 300% (was 200%). Moved Symbolize to the right column, under the Fill Area panel. Density now has a slider and moved from the Layout panel to below Size
- v2.0.0 (2026-10-27) Several objects to tile can now be selected (Order in the Tiles panel: In Order or Random). Added Brick (Vertical) and Hexagonal to the layouts and renamed Brick pattern to Brick (Horizontal). Added Density to Random, and Reshuffle draws a new arrangement, jitter and order. Added Center (on by default), Keep Overflow and Mask, and Angle. Spacing is split into Spacing X and Spacing Y, which the link icon can keep linked. Added the Tiles panel (Rotate Each, Flip Every Other 180°, Jitter with Rotation and Size, and Order). Tiles shows how many tiles will be placed, and Preview can be turned on and off. Added sliders linked to the number fields and rearranged the dialog (number settings in one column; below them the Layout panel and Symbolize on the left, the Tiles and Fill Area panels on the right; Preview to the left of the buttons). Renamed the Larger Object panel to Fill Area and ordered it Delete, Make Guides, Opacity
- v1.8.0 (2026-10-27) Replaced the Brick pattern checkbox with layout radio buttons: Grid, Brick pattern and Random. Added Random, which places tiles at random positions without overlapping, keeping the spacing. The arrangement changes each time the dialog opens and stays the same between the preview and the result while it is open. Adjust X and Adjust Y are dimmed because Random does not use them. The dialog now has two columns, with the Layout and Larger Object panels on the right
- v1.7.0 (2026-10-27) The settings confirmed with OK (scale, spacing, offset, brick pattern, symbolize) are remembered until Illustrator quits and used as the defaults next time. When the larger object is a group, tiles now fill the shape of each object in the group instead of the group's bounding box (a clipping group uses its clipping path). When it is text, tiles fill the glyphs of an outlined duplicate, which is then deleted. Added the Larger Object panel to choose what happens to the larger object after tiling: Opacity (100% by default), Delete or Make Guides. Added Adjust X and Adjust Y to nudge the tile positions (also remembered). Whether a tile fits is now worked out before it is duplicated, which makes the preview and the run faster. Edges are hidden while the script runs and restored afterwards
- v1.6.0 (2026-10-27) Added Scale. Replaced Margin with Offset (same sign as Offset Path: negative shrinks inward, positive lets tiles stick out; distance from slanted edges and curves is now measured exactly). Removed Grid count (columns and rows); tiles now always fill as much of the shape as fits. Brick pattern and Symbolize are now on by default
- v1.5.12 (2026-10-27) Renamed the script from PatternFill to TileInsideShape. Tiles that stick out of the path are now removed for any shape, not only rectangles and ellipses (previously a tile was kept when its center was inside the bounding box)
- v1.5.11 (2026-10-16) Right-only button rows are now centered when the dialog's inner width is up to 240 px (was 200 px) (shared part update)
- v1.5.10 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.5.9 (2026-10-03) Number fields now show the unit inside the field; removed the unit beside the fields
- v1.5.8 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.5.7 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.5.6 (2026-10-01) Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.5.5 (2026-09-30) Dropped the script's own rightward shift of the dialog on first open. Button rows with only right-side buttons are now centered
- v1.5.4 (2026-09-30) Fixed an error when running with characters selected by the Type tool
- v1.5.3 (2026-09-29) Dialog opacity changed to 98%
- v1.5.2 (2026-09-28) The button row is now built with the shared part
- v1.5.1 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.5.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.4.2 (2026-09-19) Merged `fillinger-pattern.jsx`; added the grid count (columns and rows)
- v1.4.0 (2025-10-30)
