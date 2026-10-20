# Redistribute rules evenly inside a frame

[![Direct](https://img.shields.io/badge/Direct%20Link-AverageRulesInRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/AverageRulesInRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Evenly redistributes the vertical and horizontal rules inside the selected outer rectangle.
Horizontal and vertical rules can be toggled independently, and rules can be stretched to the sides of the frame.

### Features

- Vertical rules are placed at equal divisions of the frame width (count + 1), horizontal rules at equal divisions of its height
- "Distribute horizontal rules" / "Distribute vertical rules" choose the directions (a direction with no rules is dimmed)
- "Match rule lengths" snaps both ends of two-point rules to the sides of the frame (when off, rules keep their length and only move)
- Preview switch. Closing with Cancel or the Esc key restores the original layout
- Option (Alt)-click either rule checkbox to toggle both at once
- Rules inside groups and compound paths are included (locked and hidden items are skipped)

### Usage

1. Select the outer rectangle together with the rules inside it.
2. Run the script. The dialog opens with the rule directions found turned on, and the preview is shown.
3. Choose the directions and "Match rule lengths", then confirm with OK while watching the preview.

### Notes

- The closed path with the largest area in the selection is used as the outer frame.
- A rule is a thin open path at least 20% as long as the matching side of the frame, with its center inside the frame. Thin filled rectangles (closed paths) are not treated as rules.
- Rules keep their left-to-right (horizontal: top-to-bottom) order.
- "Match rule lengths" only affects rules with two anchor points; other rules are only moved.
- The initial state of "Match rule lengths" can be changed with `MATCH_LINES_TO_RECT` at the top of the script.

### note

https://note.com/dtp_tranist/n/n7e8c9ce0173e

### Update History

- v1.0.9 (2026-10-20) Fixed "Match rule lengths" having no effect (the checkbox now controls it). Fixed the preview result remaining when the dialog was closed with Esc. The dialog now opens with the rule directions found and "Match rule lengths" turned on
- v1.0.8 (2026-10-16) Right-only button rows are now centered when the dialog's inner width is up to 240 px (was 200 px) (shared part update)
- v1.0.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.0.6 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.0.5 (2026-10-01) The button row now uses the shared part. Unified the window and panel margins and spacing with the shared layout part
- v1.0.4 (2026-09-30) Fixed an error when running with characters selected by the Type tool
- v1.0.3 (2026-09-29) Dialog opacity changed to 98%
- v1.0.2 (2026-09-28) The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.0 (2026-04-29) Initial release
