# Create an offset rectangle behind the selection

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawRectangleBehindSelectedObject.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DrawRectangleBehindSelectedObject.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawRectangleBehindSelectedObject.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Generate rectangles offset from the bounding box of selected objects
- Live preview with immediate feedback; created rectangles are always sent to back
- Opacity applies to both preview and the finalized rectangle

Last updated: 2025-11-09

### Key Features

- Offset (follows current ruler units)
- Corner radius (applied via Live Effect, kept unexpanded)
- Fill/Stroke color options (K100 / White / HEX / CMYK)
- Target: Individual or as Group
- Preview on a dedicated layer (does not pollute history)
- Dialog position, opacity, and parameter persistence

### Processing Flow

1. Compute bounding box of target objects
2. Apply offset, corner radius, and fill/stroke settings to rectangle
3. Render preview to dedicated layer; finalize on OK
4. Optionally group the rectangle with the original objects

### Update History

- v1.0 (2025-08-22): Initial version
- v1.1 (2025-08-23): Added preview and color selection
- v1.2 (2025-08-23): Added type (fill/stroke), stroke width, and preset saving
- v1.3 (2025-08-28): Added dialog position, opacity, and parameter persistence
- v1.4 (2025-09-02):
- v1.5 (2025-11-09): Reviewed fill logic (HEX→CMYK when needed, disable overprint, enforce Normal)
- v1.6 (2025-11-09): Preview stabilization (debounce & cancel, before/afterRender, bump compat, immediate refresh fix)
- v1.6.2 (2026-09-27): Code cleanup. Renamed the “Fill” panel to “Color” and “Group with Text” to “Group with Objects”; added colons to field labels and tooltips. Fixed the stroke width stepping twice per arrow key, values being re-rounded on keys other than the arrows, the temporary measuring layer being left behind, and the preview remaining after closing with Esc
- v1.7.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.7.1 (2026-09-28): Replaced the Link checkbox with a link icon. The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.7.2 (2026-09-28): The button row is now built with the shared part. Settings are now saved through the shared part (stored in Folder.userData/illustrator-scripts/DrawRectangleBehindSelectedObject.json). Keyboard shortcuts now use the shared part (ignored while Cmd etc. are held). Clip groups are now measured by their mask
- v1.7.3 (2026-09-29): Dialog opacity changed to 98%
- v1.7.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.7.5 (2026-09-30): Dropped the script's own rightward shift of the dialog on first open

### Script info

- Version: v1.7.5
