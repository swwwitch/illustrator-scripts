# Reorder the stacking order by position or Z index

[![Direct](https://img.shields.io/badge/Direct%20Link-ZIndexSorter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/sort/ZIndexSorter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZIndexSorter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Reorders the stacking order of objects by position (X/Y) or by Z index
- Sorts in ascending, descending or random order

### Main Features

- The sort axis and order are chosen in a dialog
- Cancelling restores the original order
- Japanese and English localization

### Process Flow

- Collect the selected items
- Read the sort axis and order from the dialog
- Sort, then restack the items in Z-index order

### Notes

- Verified on Illustrator 2025 and later

### Update History

- v1.0 (20250806): Initial version
- v1.0.2 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.0.3 (20260928): The button row is now built with the shared part
- v1.0.4 (20260929): Dialog opacity changed to 98%
- v1.0.5 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.0.6 (20260930): "Original Stacking Order" no longer flips on every click: Ascending keeps the order, Descending reverses it, Random shuffles it. Fixed Cancel restoring the order reversed, and an error when running with no document open. Position sorting is now reliable. Revised labels and tooltips
- v1.0.7 (20260930): Cancel (including Esc and the close box) now puts every object back exactly where it was, including its position relative to unselected objects and its layer. Added keyboard shortcuts for the radio buttons (Z/X/Y, A/D/R). Dropped the script's own rightward shift on first open (the dialog still moves sideways to avoid the selection). Panel margins now follow the shared settings. Button rows with only right-side buttons are now centered

### Script info

- Version: v1.0.7
