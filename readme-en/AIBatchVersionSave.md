# AIBatchVersionSave

[![Direct](https://img.shields.io/badge/Direct%20Link-AIBatchVersionSave.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/AIBatchVersionSave.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that batch-saves the .ai / .svg files in a folder using a selected Illustrator version.
- .svg files are opened in Illustrator and saved as .ai in the chosen version.
- "Overwrite mode" and "Custom" mode switch between in-place overwriting and suffix-renamed copies.
- During processing Illustrator's interaction dialogs (color profile, font substitution, overwrite confirmation, …) are suppressed and a progress window reports the status.

### Main Features

- Save-version dropdown listing only the versions supported by the running Illustrator (CC and later are consolidated into v17)
- "Overwrite mode": saves each file in the source folder as .ai under the same base name (.ai inputs are overwritten, .svg inputs produce a same-named .ai). Destination and filename-append options are disabled
- "Custom": destination folder and filename-append options are fully configurable
- "Folders" panel: pick folders with [Source...] and [Destination...]. "Use same folder as source" (default ON) lets you skip picking a destination
- Path display toggles: "Full path" and "Shorten Dropbox path" (full-path toggle is disabled while Dropbox shortening is ON; tooltips always show the full path)
- "Targets" panel: .ai / .svg checkboxes. Extensions with no matching files in the source folder are automatically dimmed and turned off
- "Save Settings" panel: use Illustrator's built-in [Converted] suffix (the preference is restored after processing), and create a PDF-compatible file
- "Filename Adjustments" panel: append the save-version string (v17 etc., updated automatically when the version changes), append custom text, and choose the separator (`-` / `_`, default `_`)
- Half-width / full-width spaces and tabs in filenames are replaced with the separator, and the output extension is always normalized to lowercase .ai
- Overwrite-safety confirmations: in Custom mode when .ai inputs exist, source and destination match, and nothing is appended; and in Overwrite mode when .svg files are part of the batch
- The progress window shows the current filename and the count (n / total)
- Tooltips on every option, with automatic Japanese / English UI

### Workflow

1. Set the mode, folders, target extensions, save version and filename adjustments in the settings dialog
2. Collect the matching files from the source folder and show the overwrite confirmations when applicable
3. Temporarily change the [Converted] preference, suppress interaction dialogs, then open and saveAs each file in turn
4. Close the progress window, restore the preference and the user interaction level, and report completion

### Not Supported

- Running Illustrator versions that expose no selectable save version
- Neither .ai nor .svg checked in the "Targets" panel
- Source folders that contain no matching files
- Files with extensions other than .ai / .svg, and files inside subfolders

### note

- Original idea: Kuro-san (VoostOn) https://vooston.web.fc2.com/dtp/dtp_a010.html

### Update History

- v1.5.0: Current version
