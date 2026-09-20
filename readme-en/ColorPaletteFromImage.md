# Build a color palette from an image

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorPaletteFromImage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ColorPaletteFromImage.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorPaletteFromImage.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script extracts representative colors from a selected placed image, raster image, vector art, or text, then draws 16 / 11 / 8 / 5 color palettes as squares below the source object.

With nothing selected, it builds a palette from the swatches currently selected in the Swatches panel.

The 5-color rows can also be registered as swatch groups.

## Main features

- Extracts colors from placed and raster images by tracing and expanding them
- Extracts colors straight from vector art without rasterizing (gradient stops are picked up individually)
- Switches automatically to rasterize-and-trace extraction when a fill cannot be read, such as a freeform gradient
- Duplicates text and converts it to outlines before extracting colors
- Outputs 16, 11, 8, and 5 color palettes fitted to the width of the source object
- Picks representative colors with an area-weighted max-distance method (larger areas are favored)
- Sorts colors into a gradient-like order with a nearest-neighbor walk
- Shows HEX under the 5-color row and CMYK under the CMYK-rounded row
- Live preview while the dialog is open
- Builds a palette from the swatch selection
- Registers the 5-color and CMYK-rounded rows as swatch groups
- Japanese and English UI

## Usage

1. Select the objects to extract colors from (leave nothing selected to use the swatch selection instead).
2. Run `ColorPaletteFromImage.jsx`.
3. Pick an image trace preset from the "Built-in" or "Custom" list when the selection contains a raster or placed image ("[6 Colors]" is preselected).
4. Choose the rows and the color info to output, then click OK.

## Dialog options

| Option | Description |
| --- | --- |
| All rows / 5-color rows only | "5-color rows only" skips the 16, 11, and 8 color rows ("5-color rows only" is the default) |
| Rows to Output | Selects the rows to output (16 / 11 / 8 / 5 / 5 with CMYK rounding) |
| Color Info | Shows HEX under the 5-color row and CMYK under the CMYK-rounded row |
| Fit View | When checked, refits the view to the source object and the preview every time a setting changes (on by default) |
| Reselect | Starts over from the image trace preset selection |

The buttons sit in a single row at the bottom of the dialog: "Reselect" on the left, "Cancel" and "OK" on the right.

## Workflow

1. Create a temporary work layer and duplicate the selected objects onto it. A clip group is duplicated as well, and it is the duplicate that gets rasterized to the bounds of its clipping path, so the original is left alone.
2. Trace and expand raster/placed images; read fill colors with their areas straight from vector art. Text is outlined first. When the vector art contains a fill that cannot be read (a freeform gradient and the like), the duplicate on the work layer is rasterized and traced instead.
3. Merge identical colors, summing their areas, then narrow the palette down in stages: 16 → 11 → 8 → 5. This selection does not depend on the output options, so it runs once per object and is reused for every preview redraw.
4. Show the dialog as soon as the first colors are available, refreshing the preview as options change. The work layer is hidden while the dialog is open, so the traced duplicate never covers the original.
5. On OK, register the swatch groups and draw the palettes. The work layer is removed afterwards.

## Scope

| | Objects |
| --- | --- |
| Handled | Placed images, raster images, paths, compound paths, groups, clip groups, text frames, swatch selection |
| Not handled | Locked objects, hidden objects |

## Notes

- Only the 5-color and CMYK-rounded rows are registered as swatch groups. The 16, 11, and 8 color rows are drawn but not registered.
- Swatch registration always creates new swatches. When a swatch of the same name already exists, a number is appended (for example "C=0 M=100 Y=100 K=0 2"). Existing swatches are never modified or moved.
- Swatch group names avoid collisions the same way: running the script twice on the same image produces "sample.psd - 5 Colors 2" and so on.
- CMYK labels are available only in a CMYK document. The checkbox is disabled in an RGB document.
- CMYK rounding snaps to 5% steps. The boundary between 0 and 5 is 2.5, and between 5 and 10 is 7.5. Only the CMYK-rounded row is rounded: a palette built from the swatch selection labels each square with the swatch's own values.
- The 11-color row excludes near-white and near-black colors so that mid-tones come through. The thresholds are the `NEAR_WHITE_*` / `NEAR_BLACK_*` variables at the top of the script.
- Representative colors are weighted by area (`pow(area / average area, 0.75)`). A color covering four times the average area carries roughly 2.83 times the weight.
- Clip groups are duplicated onto the work layer, and the duplicate is rasterized at 300 ppi on a white background. The original stays as it is, so cancelling the dialog leaves the artwork untouched. The settings are the `RASTERIZE_*` variables at the top of the script.
- The palette cannot be drawn on a locked or hidden layer. Objects on one are reported with an alert and skipped; the rest of the selection is still processed.
- A temporary layer named `__workLayer__` is created during the run and removed when it finishes.
- ExtendScript cannot read the colors of a freeform gradient. On Illustrator 30.7 such a fill comes back as `GrayColor` (gray=0), with nothing to tell it apart from a real gray fill. A fill is therefore treated as "not reliably read" when any of the following holds:
  - Reading the fill threw an error
  - It is a gradient but not a single color stop could be read
  - It is a `GrayColor`
- Whenever even one such fill is found, the duplicate on the work layer is rasterized (300 ppi) and traced, and the colors are read from that result instead. The original objects are left untouched. The preset dialog is not shown for a vector selection, so this trace uses the "[6 Colors]" preset.
- As a side effect, grayscale artwork also goes through the rasterize path. The traced colors are gray either way, so the palette is nearly the same; only the processing time goes up.
- Vector art and text are treated as a single palette even when several objects are selected. The palette is anchored to the top-left of the combined bounds.
- The image trace preset dialog appears only when the selection contains a raster or placed image. It opens with "[6 Colors]" preselected, looking through the custom list as well when the built-in list has no match, and falling back to the last custom preset when neither does. Change the preferred preset with `DEFAULT_TRACING_PRESETS` at the top of the script (preset names are localized, so brackets, spaces, letter case, and full-width digits are normalized away before the candidates are matched exactly, in order — substring matching is avoided so that "[16 Colors]" is never picked up by mistake).
- Labels use Myriad Pro, falling back to the default font when it is not available.
- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.

## Changelog

- v1.7.5 (20260920): Bug fixes, plus a pass over the workflow and the code. Fixed clipped groups being rasterized in place: the original was replaced by a flat raster before the dialog was ever shown, so cancelling left the artwork destroyed. The rasterizing now happens on a duplicate on the work layer, and the original is never touched. Fixed the traced duplicate (the image reduced to 6 colors) covering the original while the options dialog was open; the work layer is now hidden for as long as the dialog is up. Fixed the CMYK labels of a palette built from the swatch selection showing values that differ from the color actually drawn, because the 5% rounding was applied to every label instead of only to the CMYK-rounded row. Fixed a second run on the same image colliding with the swatch group name it created the first time, which threw and ended the run with no palette drawn and no message: group names now get a number appended, and a failed registration no longer stops the drawing. Fixed clicking the already-selected "5-color rows only" radio wiping the saved state of the 16, 11 and 8 color rows, and fixed the HEX / CMYK choices made in that mode being reverted on the way back to "All rows". Fixed `GrayColor` fills having no CMYK conversion, which left them unrounded in the CMYK-rounded row and labeled with a hex value. Fixed the group name staying "5 Colors" when only the CMYK-rounded row was requested. Fixed objects on a locked or hidden layer being skipped silently; they are now reported before any work starts. Dropped the full preview that was drawn and then thrown away before the dialog opened, and the row selection now runs once per object instead of twice, and is reused for every preview redraw rather than being recomputed on each checkbox click. The 11-color candidates now reuse the RGB values already computed instead of converting each color again. Split the output options dialog into control building, the button row, and state handling, and split the area weighting and the farthest-color search out of the selection routine. Consolidated the name-collision handling into one function, tidied the color-matching key and the entry structure, and removed the branches nothing reached (cascade off, empty-row drawing, modulo indexing, and the `app.selection` assignment). Added tooltips to "8 Colors" and "5 Colors"
- v1.7.4 (20260815): Fixed colors not being extracted from freeform gradients. On Illustrator 30.7 such a fill comes back as `GrayColor` (gray=0) rather than `GradientColor`, so it threw no error, was never treated as a gradient, and contributed a single meaningless gray. `GrayColor` fills, gradients with no readable stops, and fills that throw on access are now all flagged as "not reliably read", and if there is even one, the duplicate on the work layer is rasterized and traced as a fallback. Also fixed a fill that throws on access dropping the palette for the whole vector selection. Fixed swatch registration moving existing same-named swatches into the palette group. A swatch belongs to only one group in Illustrator, so `addSwatch()` pulled the user's swatch out of its own group, making user-defined swatches look like they had vanished. The generated names (`C=0 M=100 Y=100 K=0` form) match the default swatch names, so collisions were likely. Swatches are now always created fresh instead of being reused. Also reorganized the whole script to match the house rules. Wrapped everything in an IIFE, rewrote the overview and basic-info blocks, categorized LABELS behind a single `getLabel()`, and added JSDoc to every function. Renamed variables, panels, groups, and functions to self-explanatory names. Consolidated the duplicated 5% rounding, hex conversion, font assignment, nearest-neighbor sort, and cascade selection, and split the top-level main flow into functions. Removed the unused `infoMode` field, the always-true `preview` field, and the dead swatch-group input branch. Added tooltips to the main checkboxes, buttons, and preset lists. Reworded the UI to match what it actually controls ("Color Counts" → "Rows to Output", "Info for all rows / 5-color rows only" → "All rows / 5-color rows only", "Retry" → "Reselect") and added "Built-in" / "Custom" headers above the preset lists. The swatch-selection label is now a single two-line frame (CMYK above HEX) instead of two frames, and the near-white / near-black test now works for LabColor as well. Preset names are now normalized (brackets, spaces, letter case, full-width digits) before matching, and the custom list is searched when the built-in list has no match. The buttons moved into a single row at the bottom of the dialog, with "Reselect" on the left and "Cancel" / "OK" on the right, and the "All rows / 5-color rows only" radios are centered. "Fit View" became a checkbox (on by default) that refits the view every time a setting changes, instead of a button that fits it once
- v1.7.3 (20260417): Minor fixes
- v1.7.2 (20260417): Minor fixes
- v1.7.1 (20260320): Minor fixes
- v1.7 (20260306): Added image trace preset selection and a Retry option to reselect the preset after building a palette
- v1.5 (20260305): Changed representative color selection to take area into account
- v1.4 (20260305): Skipped rasterize and image trace for vector art
- v1.2 (20260305): Initial version
