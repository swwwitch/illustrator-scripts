# Lower the auto-leading percentage by ten steps

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoLeadingStep--10.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/AutoLeadingStep-10.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingStep-10.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sets the auto-leading amount (%) so the selected text's displayed leading value steps up by one
integer (e.g. 26.124 → 27). A sibling of AutoLeadingCalc.jsx — no dialog is shown; it applies to
the selection in place, including text inside groups and range selections in text-edit mode. For
each paragraph the current leading is converted to the document's text unit, the next integer is
chosen, and the auto-leading amount that yields that integer leading is applied.

### Script info

- Version: v1.0.3

### Update History

- v1.0.2 (2026-09-28): With the type unit set to feet/inches, one step now equals one unit (a foot) instead of one inch
- v1.0.3 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
