# Batch rename and reorder artboards, symbols and layers

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartRenamer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SmartRenamer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenamer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

    - Batch rename and reorder Illustrator artboards / symbols / layers / graphic styles by combining prefix, suffix, name source, and find/replace.
    - As settings change in the dialog, results are updated automatically for preview.

### Item Type (top of dialog)

    - Switch between Artboard / Symbol / Layer
    - On switch, the right-side list and Filter / Name Source panels are rebuilt for that type
    - In Symbol / Layer mode the Name Source panel remains visible, and renaming can use "Original Name" or custom text plus prefix / suffix / find-replace

### Rename conditions (left column)

    - Prefix / Suffix: sequence tokens (`{#1}` → 1,2,3 ; `{#01}` → 01,02,03 with zero padding), `#FN` (file name), `#DT` (date), separators (`-` `_`)
    - Name Source: "Original Name", "Frontmost Text" (Artboard mode only; recursively scans layers/groups and uses the first visible unlocked TextFrame found inside the artboard), or "Custom" (free text)
    - Find / Replace: find and replace strings. The Regex checkbox switches between RegExp mode (default ON) and literal (escaped) mode. The replace field also has token buttons (`#`→`{#1}`, `##`→`{#01}`, `-`, `_`, `x` to clear). The find field also includes shortcut buttons for `\d`, `\d+`, and `.+` (match-all)
    - Automatically appends "_1", "_2", etc. only on name collision (skipped when prefix / suffix / replace contains a sequence token, since results are assumed unique)

### Filter (top-right panel)

    - Row 1: "All" or "Range" (`1-3,5` form). Two-way linked with the Reorder / Rename checkboxes
    - Row 2: "Filter by search" checkbox + text. When ON, ignores the All/Range selection and checks only items whose current name contains the text
    - The filter UI is grouped into its own dedicated panel at the top of the right column

### Reorder / Rename (right column, bottom)

    - Lists each item as "Current Name -> New Name". After Refresh, committed current names are reloaded as the right-column baseline names, and unchanged OK execution avoids double-apply
    - Checked rows allow manual override of the new name (manual edits stay sticky)
    - Move checked rows with Top / Up / Down / Bottom. Artboards swap rects, layers use `.move()` (symbols / graphic styles cannot be reordered)
    - Option-click a checkbox:
      - When no rows are checked: turn all rows ON
      - When some rows are checked: bulk toggle to the new value
      - When all rows are checked: isolate (only the clicked row stays ON)

### Buttons

    - "Refresh" immediately commits current settings, reorder state, and manual-edited names. Symbols and layers are also updated immediately before pressing OK
    - Cancelling restores pre-dialog names for all three item types

### Update History

    - v1.0 (20250509): Initial version
    - v1.5.2 (20260508): Refresh now commits right-column manual name edits and reorder changes
    - v1.5.3 (20260508): On name collisions, the first occurrence now starts at "_1" (non-duplicates remain untouched)
    - v1.6.0 (20260511): Added item type switching (artboard / symbol / layer), dedicated filter panel, "Original Name" mode, search filter, regex shortcut buttons, find / replace (regex-capable), `{#N}` sequence tokens, Option-click isolate, post-Refresh re-baselining, immediate symbol/layer Refresh updates, and double-apply prevention when pressing OK without changes after Refresh
    - v1.6.2 (20260923): Row "New Name" edits now count as manual edits even when Refresh / OK is pressed without leaving the field. Fixed: manual checkbox changes under the search filter being overwritten by Refresh, the Range field re-enabling while the search filter is on, and `{#01}` sequences truncating to "00" from item 100. Reorder buttons are now disabled for symbols / graphic styles, which cannot be reordered (the reorder silently failed and other items got renamed). The list now shows 12 rows with a scrollbar (rows beyond the panel height used to be unreachable). OK with "Original Name" and no input is now treated as no input (duplicate names no longer get "_1") and warns before the dialog closes. Fixed: "Frontmost Text" being assigned to another artboard when artboards overlap, and temporary names being left behind when an artboard rect fails to restore

### Script info

- Version: v1.6.2
