# Grid AI/PDF files into a thumbnail contact sheet

[![Direct](https://img.shields.io/badge/Direct%20Link-SlideCollage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SlideCollage.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Places selected .ai / .pdf files (PDFs by page) on the active document in a grid to build a portfolio-style thumbnail sheet.

- Import: place by artboard number (for example `1-20` or `1,3,5`)
- Item: choose the PDF crop box (art / trim / bleed / bounding) and apply a corner radius (converted to points)
  - Each item is clipped into a group by a same-sized rectangle, and the corner radius (a live effect) is applied to that clip group
  - Spreads (landscape pages) can be split into left and right halves; the side for even pages is set from the PDF binding direction
- Grid: set direction (horizontal / vertical / random), column count and spacing
- Even columns: a distribution mode adds one extra slot to even columns, and their vertical offset can be tuned separately
- Layout: scale (an extra multiplier on top of the auto-fit result), rotation (rotates the whole sheet and centers it on the artboard) and position offset (horizontal / vertical)
- Mask: clips inside the margins on OK, with an adjustable mask corner radius applied to the clip group
- Artboard: an optional background colour (entered as HEX) can be added, and it is excluded from the mask

Most changes are reflected in a real-time preview; changes to Range and Total take effect with [Load].
Numeric fields step with Up/Down, Shift and Option.

### Original idea

Slide Collage - Portfolio Layout Generator -

https://slide-collage.vercel.app/

### Article

https://note.com/dtp_tranist/n/n9f8c7370f4e5

### Script info

- Version: v1.7.1

### Update History

- v1.7.1 (2026-09-29): Split spreads are now laid out in a grid sized to a single page; the column limit is now 30 instead of 10
- v1.7.1 (2026-09-29): Removed the zoom slider and Light Preview; round corners are no longer previewed and are applied on OK. Moved the Load button below Total, and added stepper buttons to Total. Moved Margin to the right of Mask and lined up Round Mask Corners with it. Checkboxes followed by a value (Round Corners, Shift, Rotate, Background, Offset, Round Mask Corners) now end with a colon too
- v1.7.1 (2026-09-29): Added Fit to Window (zooms so the artboard fills the given share of the window; Cancel restores the view). Moved the Artboard & Mask panel to the right column
- v1.7.1 (2026-09-29): Turning on Shift now fills in the offset automatically (half of an item's height plus the gap)
- v1.7.1 (2026-09-29): Fixed the source document being closed without saving when it was already open and its page count could not be read from the file
- v1.7.1 (2026-09-29): Fixed the color picker's color not reaching the background in CMYK documents. Label widths are now measured so labels such as "Scale:" are not cut off. Added a tooltip to Round Corners
- v1.7.0 (2026-09-29): Added splitting spreads into left and right halves (ported from PDFAIImporter)
- v1.7.0 (2026-09-29): Fixed the mask, mask corner radius and corner radius not being applied when OK was clicked after loading
- v1.7.0 (2026-09-29): Fixed the position offset not showing in the preview without rotation; Total now sets the number of placed items
- v1.7.0 (2026-09-29): Changing the crop box now reloads the preview; decimals can be typed into Gap and the other fields again
- v1.7.0 (2026-09-29): In English, Reset and Light Preview are no longer shown in Japanese; field labels get colons and tooltips were added
- v1.7.0 (2026-09-29): Reorganized the code (shared placement and grid math, structured labels, clearer names)
- v1.6.1 (2026-09-28): The button row is now built with the shared part
- v1.6.1 (2026-09-28): In English, removed the extra space after field-label colons
- v1.6.1 (2026-09-28): Fixed an error when closing with OK or Cancel, and pages not repeating when there are fewer source pages than requested
- v1.6.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.5.2 (2026-09-25): Fixed the crop box values so Trim, Bleed and Art place the box you choose, and corrected the English names of Crop and Trim
