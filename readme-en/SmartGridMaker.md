# SmartGridMaker.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-SmartGridMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/SmartGridMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script builds a frame and grid in one pass, based on either a selected rectangle or the artboard.

Outer area (edge scale, line caps, rounded corners), title area, inner area offsets, column and row divisions with spacing, fills, dividers, line types, and a bleed-aware frame are all configured in a single dialog with preview.

## Main features

- Works from a selected rectangle or from the artboard
- Margin settings (top, bottom, left, right, with link) when based on the artboard
- Outer area edge scale, line caps (Butt, Round, Project), and rounded corners
- Title area (Top, Bottom, Left, Right position, size, fill, line, edge scale)
- Inner area offsets (top, bottom, left, right, with link)
- Column and row counts with spacing, plus dividers
- Solid, dash, and dot-dash line types
- Bleed-aware frame (artboard-based only)
- Pan and zoom controls plus view commands (Fit Artboard, Actual Size, Fit All) in the Display tab
- Dialog with preview
- Values are entered in Illustrator's ruler unit
- Dialog settings are restored on the next run (reset when Illustrator restarts)
- Japanese and English UI

## Usage

1. Select a rectangle to use as the base. To use the artboard instead, run the script with nothing selected.
2. Run `SmartGridMaker.jsx`.
3. Configure the settings in the dialog tabs.
   - Margin: margins from the artboard (artboard-based only)
   - Outer Area: edge scale, line caps, rounded corners, title area, frame
   - Inner Area: offsets, columns and rows, spacing, fill, dividers, line type
   - Display: pan and zoom, view commands
4. Adjust the values while checking the preview.
5. Click OK to generate the result.

## Tabs and settings

| Tab | Settings |
| --- | --- |
| Margin | Top, bottom, left, and right margins, link |
| Outer Area | Edge scale, line caps, rounded corners, title area, frame (bleed) |
| Inner Area | Offsets (top, bottom, left, right, link), column and row counts, spacing, fill, dividers, line type |
| Display | Zoom, horizontal and vertical pan, view commands |

## Notes

- When the script starts from a selected rectangle, the frame panel is disabled and the margin panel is hidden; both apply to artboard-based runs only.
- Bleed applies to the frame only.
- The title area uses the outer area's rounded-corner value on the two corners that match its position, and never rounds the inner area.
- Edge scale is dimmed while the title area's Line option is off.
- When the title area is enabled, its default size is one fifth of the outer area's height.
- Values follow Illustrator's ruler unit setting.
- Dialog settings persist only while Illustrator is running.

## Article

[Generate a frame and grid with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n2b01f896c423)

## Changelog

- v1.4.1 (2026-02-24): Improved the stability of rounded-corner handling (Error 23)
- v1.4.0 (2026-02-24): Added rounded corners for the title area and a Display panel
- v1.3.0 (2026-02-24): Split margins into four sides with link support and refactored UI construction
- v1.2.0 (2026-02-24): Changed the inner area offset UI to a three-column layout and added rounded corners to the outer area
- v1.1.0 (2026-02-24): Separated preview and generation and added session restore
- v1.0.0 (2026-02-24): Initial version
