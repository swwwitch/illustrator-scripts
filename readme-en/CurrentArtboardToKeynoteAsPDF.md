# Insert the active artboard into the current Keynote slide

[![Direct](https://img.shields.io/badge/Direct%20Link-CurrentArtboardToKeynoteAsPDF.applescript-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/AppleScript/CurrentArtboardToKeynoteAsPDF.applescript)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CurrentArtboardToKeynoteAsPDF.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An AppleScript that exports the active artboard as PDF and overlays it as an image on the slide currently shown in Keynote
- The no-dialog version: it does what [ArtboardsToKeynoteAsPDF](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardsToKeynoteAsPDF.md) does by default, sized to fit the slide

### Main Features

- Exports the whole artboard as PDF and centers it at the largest size that fits the slide
- Shows the target slide in Keynote after inserting

### Usage

1. Show the target slide in Keynote.
2. Make the artboard you want to insert active in Illustrator.
3. Run the script.

If no document is open in Keynote, a new one is created.

### Notes

- An artboard with no objects is not inserted, and a warning is shown.
- The target application is "Keynote Creator Studio".
- On first run, macOS asks for permission (Automation) to control Illustrator and Keynote.

### Update History

- v1.0.0 (2026-10-21): Initial release
- v1.0.1 (2026-10-22): Fixed the image landing on an earlier slide than the one shown when the document has skipped slides

### Script info

- Version: v1.0.1
