# Slice an image or artwork into grid or puzzle pieces and mask them

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSliceWithPuzzlify.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/SmartSliceWithPuzzlify.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSliceWithPuzzlify.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Slices the selected image or artwork into grid cells or jigsaw pieces and masks each piece.

![ss-2092-1082-72-20250608-151834](https://github.com/user-attachments/assets/b57117df-bfbc-49d7-b648-c6a89984999c)

### Main Features

- Two methods: Grid and Puzzle. Puzzle offers Traditional (regularly placed tabs) and Random (random tab directions)
- Enter a piece count and the columns and rows are calculated from the selection's aspect ratio so that each piece is roughly square
- Placed images and symbols are duplicated for each piece and masked with a clipping group
- Embedded images and vector artwork (paths, groups, compound paths) are converted to a symbol before slicing. Multiple selected objects are grouped into one symbol, keeping their stacking order
- Fine-tune the result with offset, overlap, scatter, stroke and rounded corners
- Japanese and English UI

### How to Use

1. Select the image or artwork to slice
2. Run the script and choose the method and options in the dialog
3. Click OK. A progress bar is shown while slicing, and the dialog closes when it finishes

### Options

| Option | Method | Description |
|---|---|---|
| Method | — | Grid / Puzzle |
| Pieces | Puzzle | Approximate number of pieces. With one object selected, the columns and rows follow from its aspect ratio |
| Columns / Rows | Both | Pieces across and down. Enter 0 in one of them to derive it from the other and the aspect ratio |
| Shape | Puzzle | Traditional / Random |
| Offset | Puzzle | Offsets the outline of each piece. A negative value shrinks it inward |
| Overlap | Grid | How far neighbouring pieces overlap, to hide the seams (never beyond the original bounds) |
| Scatter | Puzzle | Moves each piece by a random, center-biased amount. The value is the maximum distance |
| Add stroke | Both | Adds a stroke to each piece |
| Round corners | Grid | Applies the Round Corners effect to each piece. The value is the radius |

- Distances are entered in the ruler units
- In number fields, the stepper buttons on the left and the Up/Down keys move the value to the next whole number (1.5 → 2), with Shift to the next multiple of 10, and with Option by 0.1 (pieces, columns and rows take whole numbers only)
- Switching the method resets the options to their defaults

### Notes

- The original object is deleted after slicing
- Embedded images, vector artwork and multiple selections are converted to symbols, so new symbols appear in the Symbols panel
- For any other object (such as text), only the mask paths are created and the original is kept
- Offset applies the Offset Path effect and then expands the appearance

### Article

[【Illustrator】画像だけを選択して、パズルのピース作成からマスクまでを一括で行うスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n89f63325c0bc?magazine_key=mebd7eab21ea5) (Japanese)

### Original / Acknowledgements

Originally created by Jongware on 7-Oct-2010  
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### Update History

- v1.0.0 (20250607): Initial version
- v1.0.1 (20250608): Added symbol and vector artwork support
- v1.0.2 (20250609): Added grid shape support
- v1.0.3 (20250610): Added offset feature and unit code support
- v1.5.2 (20260927): "Add stroke" now works in Puzzle mode too; revised the dialog title, field labels and tooltips; options disabled for the current method (such as Scatter in Grid mode) no longer take effect; the dialog stays open when the column/row counts cannot be sliced
- v1.6.0 (20260927): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.6.1 (20260928): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.6.2 (20260928): The button row is now built with the shared part
- v1.6.3 (20260929): Dialog opacity changed to 98%
- v1.6.4 (20260930): Fixed an error when running with characters selected by the Type tool
- v1.6.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
