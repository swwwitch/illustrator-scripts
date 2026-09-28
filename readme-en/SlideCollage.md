# Grid AI/PDF files into a thumbnail contact sheet

[![Direct](https://img.shields.io/badge/Direct%20Link-SlideCollage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SlideCollage.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lays out the artboards (or PDF pages) of the .ai / .pdf file you choose in a grid on the active document, building a portfolio-style thumbnail sheet. The size that fits the artboard is calculated automatically.

### Features

- Load: set the Range (for example `1-20` or `1,3,5`) and the Total; when Total is larger than the range, the range repeats
- Items: PDF crop box (Art / Crop / Trim / Bleed) and round corners
- Spreads: landscape pages can be split into left and right halves and laid out at single-page size; the side for even pages (right / left) is set from the PDF binding direction
- Grid: flow (horizontal / vertical / random), columns (1–30) and gap
- Even columns: +1 Slot adds one slot to even columns; Shift moves even columns up or down (turning it on fills in half a cell automatically)
- Layout: scale (a multiplier on top of the fitted size), rotation (rotates the whole layout and centers it on the artboard), and horizontal / vertical offsets
- Artboard & Mask: background color (click the swatch for Illustrator's standard color picker, or enter HEX), mask (on OK, clips inside the margin) and round mask corners
- Fit to Window: zooms so the artboard fills the given share of the window

### How to use

1. Open a document and run the script. If a placed PDF / AI image is selected, its file becomes the source file
2. Choose a .ai / .pdf with [Choose File...]; Range and Total are filled in
3. Click [Load] to preview
4. Adjust the settings (the preview follows) and click [OK]

### Notes

- Changes to Range and Total take effect with [Load]; changing the crop box or the spread split reloads automatically
- Round corners, the mask and round mask corners are not previewed; they are applied on [OK]
- Number fields step with the steppers or the Up/Down keys: to the next whole number (1.5 → 2), Shift to the next multiple of ten, Option by 0.1
- [Reset] restores the settings other than the file and range (4 columns, rotation off, Shift off; the horizontal offset is set so the grid is centered)
- [Cancel] removes the placed items and restores the view
- The dialog reopens where it was last closed and moves aside when it would cover the selected objects

### Original idea

Slide Collage - Portfolio Layout Generator -

https://slide-collage.vercel.app/

### Article

https://note.com/dtp_tranist/n/n9f8c7370f4e5

### Script info

- Version: v1.7.2

### Update History

- v1.7.2 (2026-09-29): The dialog now reopens where it was last closed and moves aside when it would cover the selected objects; its opacity is now 97%
- v1.7.2 (2026-09-29): Added stepper buttons to Total. Moved Margin to the right of Mask and lined up Round Mask Corners with it. Checkboxes followed by a value (Round Corners, Shift, Rotate, Background, Offset, Round Mask Corners) now end with a colon too
- v1.7.2 (2026-09-29): Fixed the source document being closed without saving when it was already open and its page count could not be read from the file. Fixed the script stopping when run with text selected, and the background sometimes coming to the front
- v1.7.1 (2026-09-29): Split spreads are now laid out in a grid sized to a single page; the column limit is now 30 instead of 10
- v1.7.1 (2026-09-29): Removed the zoom slider and Light Preview; round corners are no longer previewed and are applied on OK. Moved the Load button below Total
- v1.7.1 (2026-09-29): Added Fit to Window (zooms so the artboard fills the given share of the window; Cancel restores the view). Moved the Artboard & Mask panel to the right column
- v1.7.1 (2026-09-29): Turning on Shift now fills in the offset automatically (half of an item's height plus the gap)
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
