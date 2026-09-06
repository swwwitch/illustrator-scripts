# Run "Select > Same > Appearance" from several reference objects at once

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectSameAppearanceMulti.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectSameAppearanceMulti.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameAppearanceMulti.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Runs Select > Same > Appearance once for each selected object and reselects everything the passes turn up. The built-in command accepts only one reference object; this script processes several references in a single run.

### Features

- Uses every selected object, one at a time, as the reference for Same > Appearance
- Combines the results of all passes and selects them together at the end
- No dialog — just select and run
- Japanese / English UI (alert messages only)

### Usage

1. Select the objects you want to use as references.
2. Run the script.
3. Every object matching any of the references is left selected.

For example, selecting one red-stroked object and one blue-filled object selects all red-stroked and all blue-filled objects at once.

### Notes

- If no document is open, or nothing is selected, the script shows an alert and exits.
- The reference objects themselves are included in the result.
- Matching is whatever Illustrator's Same > Appearance considers a match; the script does not change that behavior.
- Locked or hidden objects cannot be selected, so they never appear in the result.
- One menu command runs per reference object, so a large selection takes longer.
- The selection changes while the script runs. When it finishes, the result replaces the selection — the original selection is not restored.

### Update History

- v1.0.0 (20260906) : Initial release
