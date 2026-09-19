# Generate a border and grid in one pass

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartGridMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/SmartGridMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Builds a border and a grid in one pass, based on a selected rectangle or on the artboard.

The outer frame (edge extension, line caps, rounded corners), the title area, the inner-area offsets and column and row divisions, the line types, and a bleed-aware frame are all set in a single dialog with a live preview.

### Main features

- Works from a selected rectangle, or from the artboard when nothing is selected
- Margins (top, bottom, left, right, with link) for artboard-based runs
- Outer frame edge extension, line caps (Butt, Round, Projecting) and rounded corners
- Title area (top, bottom, left or right position, size, fill, divider, divider extension)
- Inner-area offsets (top, bottom, left, right, with link), column and row counts with spacing, cell fills and dividers
- Divider line types (Solid, Dashed, Dotted)
- Bleed-aware frame, 3 mm (artboard-based runs only)
- Zoom and pan plus view commands in the Display tab (Fit Artboard in Window, Actual Size, Fit All in Window)
- Up and Down keys step the numeric fields (Shift for ±10 snapped to tens, Option for ±0.1)
- Tooltips on every option
- Dialog with a live preview
- Values are entered in Illustrator's ruler unit
- Dialog settings are restored on the next run (reset when Illustrator restarts)
- Japanese and English UI

### Usage

1. Select a rectangle to use as the base. To use the artboard instead, run the script with nothing selected.
2. Run `SmartGridMaker.jsx`.
3. Configure the settings in the dialog tabs.
4. Adjust the values while checking the preview.
5. Click OK to generate the result.

### Tabs and settings

| Tab | Settings |
| --- | --- |
| Artboard | Margins (top, bottom, left, right, link), frame (width, bleed, rounded corners) |
| Outer | Keep outer frame, rounded corners, edge extension, line caps, title area |
| Inner Area | Offsets (top, bottom, left, right, link), column and row counts, spacing, fill, dividers, line type |
| Display | Zoom, horizontal and vertical pan, view commands |

The Artboard tab is hidden when the script starts from a selected rectangle; margins and the frame apply to artboard-based runs only.

### Notes

- Rounded corners and the edge extension cannot be combined: turning one on turns the other off.
- Any edge extension other than zero splits the outer frame into four straight lines and removes the original rectangle.
- Line caps are only available while the outer frame is split into four lines (Keep outer frame, with a non-zero edge extension).
- The title area reuses the outer frame's radius on the two corners that match its position, and never rounds the inner area.
- The default title size is one fifth of the height for a top or bottom title, and one fifth of the width for a left or right one.
- A positive divider extension shortens the divider at both ends; a negative one extends it.
- While the inner-area Fill is off, no cell fills are created, so the preview matches the result.
- Columns and rows are capped at 100.
- Bleed applies to the frame only, never to the base rectangle.
- Values follow Illustrator's ruler unit; each default is defined in millimetres and converted to the current unit.
- Cancelling drops the generated items and also restores the selected rectangle's fill and stroke as well as the view.
- On an artboard-based run, a locked or hidden active layer makes the base rectangle impossible to create, so the script reports it and stops.
- The selection is cleared after the run. Internal tags are kept in the Note field only, and the names shown in the Layers panel are cleared.
- Dialog settings persist only while Illustrator is running.

### Article

[Generate a frame and grid with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n2b01f896c423)

### Changelog

- v1.6.1 (2026-09-16): Revised the UI wording (tab names, the title area's Divider and Extend divider options, line types as Solid, Dashed and Dotted) and added tooltips to every option. Fixed 15 issues, including the preview not matching the result, short edges turning inside out with the edge extension, and the inner area disappearing for a left or right title. Reorganised the internal naming and structure
- v1.4.1 (2026-02-24): Improved the stability of rounded-corner handling (Error 23)
- v1.4.0 (2026-02-24): Added rounded corners for the title area and a Display panel
- v1.3.0 (2026-02-24): Split margins into four sides with link support and refactored UI construction
- v1.2.0 (2026-02-24): Changed the inner area offset UI to a three-column layout and added rounded corners to the outer area
- v1.1.0 (2026-02-24): Separated preview and generation and added session restore
- v1.0.0 (2026-02-24): Initial version
