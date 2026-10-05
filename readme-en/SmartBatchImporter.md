# Batch import several Illustrator files

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartBatchImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/SmartBatchImporter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBatchImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Imports open documents or Illustrator files (.ai / .svg / .eps) from a folder into one document, arranged in a grid that comes out close to square.

### Main Features

- **Source**
  - Open documents, or a folder (optionally including subfolders)
  - Filter by file format (AI / SVG / EPS) and a file-name regular expression. Counts appear after the radio buttons
- **Destination**
  - The current document (placed below its existing artboards) or a new document
  - For a new document, Settings... sets the profile (Print, Web, etc.), color mode, resolution and size (presets, mm / px)
  - Split every N files into separate new documents
- **Import options**
  - Import per artboard: imports each artboard, keeping its size and the content's position. Locked and hidden objects are included. Choose Artboard 1 only, All, or Specify (e.g. 1, 3-5)
  - Add file names as labels: adds the source file name below the imported content, on the "_label" layer
  - Include guides: excludes ruler guides and imports horizontal guides up to the artboard's shorter side and vertical guides up to its longer side, even when Lock Guides is on
  - Scale: scales the content and artboards by a percentage (stroke widths too)
  - Spacing: gap between the arranged items, in ruler units
  - After import: close the open source documents or keep them open
- Keeps the source layer structure when Paste Remembers Layers is on
- Progress bar with Cancel (items imported so far remain)
- Japanese and English UI

### Process Flow

1. Open the documents to import, or prepare a folder.
2. Run the script, choose the source, destination and options, and click OK.
3. When the import finishes, the view zooms to fit everything.

### Notes

- With After import set to Close, the source documents are closed without saving. You are asked first when any have unsaved changes
- Files opened from a folder are always closed after import
- With Import per artboard off, only visible, unlocked objects are imported

### Update History

- v1.0.0 (20250529): Initial version
- v1.0.1 (20250529): Changed folder import behavior, moved labels to "_label" layer
- v1.0.2 (20250529): Added file count display, progress bar, and cancel option
- v1.0.3 (20250529): Added progress count display (n/N)
- v1.4.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.4.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.4.2 (20260928): In English, the paste-failure message now uses a half-width colon followed by a space instead of a full-width colon
- v1.4.2 (20260928): The button row is now built with the shared part
- v1.4.2 (20260928): Clip groups are now measured by their mask (affects the position within the artboard and the cell size)
- v1.4.3 (20260929): Dialog opacity changed to 98%
- v1.4.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.4.5 (20260930): Button rows with only right-side buttons are now centered
- v1.4.6 (2026-10-01): Button rows with only right-side buttons are now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones
- v1.4.7 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.4.8 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.4.9 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.5.0 (2026-10-05):
  - Added a new document profile choice (Print, Web, etc.), Spacing, Include subfolders, and Split every N files (into separate new documents)
  - The new document's color mode, resolution and size moved to a separate dialog box opened with Settings...
  - The source layer structure is kept when Paste Remembers Layers is on
  - Guides are now imported even when Lock Guides is on. Guides are included when horizontal ones are no longer than the artboard's shorter side and vertical ones no longer than its longer side (was under half the canvas)
  - Sizes in px are now treated as 1 px = 1 pt (Full HD used to come out as 1440 × 810)
  - The destination selection is cleared before pasting
  - "After import" now defaults to "Keep open", "Import per artboard" to off, and "Include guides" to on
  - File counts moved from the panel title to after the radio buttons. Width / height / unit labels are right-aligned, and UI wording was revised
- v1.5.1 (2026-10-05): Moved "(except ruler guides)" from the Include guides checkbox to its tooltip
