# Insert artboards into Keynote slides as PDF images

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardsToKeynoteAsPDF.applescript-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/AppleScript/ArtboardsToKeynoteAsPDF.applescript)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardsToKeynoteAsPDF.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- An AppleScript that exports Illustrator artboards one at a time as PDF and inserts each into a Keynote slide as an image
- It is fast because it uses neither the clipboard nor keystrokes, and the placed images are PDF (they stay vector)

### Main Features

- Choose the artboards: current artboard only, all, or a range
- Place at a given scale (%). 100% matches the artboard size (1 pt = 1 point); the default is the largest value at which every artboard fits on a slide
- "Export the whole artboard" chooses whether the PDF covers only the artwork or the whole artboard
- In an existing document, choose between adding slides at the end and overlaying existing slides
- Images are centered on the slide

### Usage

1. Open the document in Illustrator.
2. Run the script from Script Editor or similar.
3. Set the artboards and options in the dialog box and click Insert.

If no document is open in Keynote, a new one is created and insertion starts at slide 1.

### Options

| Item | Behavior |
| --- | --- |
| Artboards | Current artboard only / All / Range. Defaults to the current artboard |
| Scale | Scale (%) relative to the size of the exported area |
| Export the whole artboard | When on, the PDF covers the whole artboard, including the margin around the artwork |
| Add slides at the end | Adds a new slide at the end for each artboard |
| Overlay existing slides | Places the images on existing slides in order from the start slide, adding slides at the end when it runs out |
| Start slide | Number of the slide where the overlay starts. Defaults to the slide being shown |

### Notes

- The start slide number is the number shown in Keynote's navigator, which does not count skipped slides. Overlaying also skips skipped slides.
- Artboards with no objects, or ones that fail to export, are not inserted; their numbers are listed at the end.
- The target application is "Keynote Creator Studio".
- On first run, macOS asks for permission (Automation) to control Illustrator and Keynote.

### Update History

- v1.0.0 (2026-10-21): Initial release
- v1.0.1 (2026-10-22): Fixed images landing on an earlier slide than specified when the document has skipped slides

### Script info

- Version: v1.0.1
