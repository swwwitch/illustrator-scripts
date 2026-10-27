# Create samples of pattern and gradient swatches

[![Direct](https://img.shields.io/badge/Direct%20Link-PatternSwatchSquares.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/PatternSwatchSquares.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternSwatchSquares.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Creates samples by filling rectangles, one per swatch, with the swatches of Illustrator's bundled swatch libraries (patterns and gradients) or of the document.
Several libraries can be chosen; one block per library is stacked vertically.

### Main Features

#### Swatches

- Choose where the swatches come from: Library or Document
- For Document, choose Selected Swatches (those selected in the Swatches panel) or All Swatches. Swatches of any kind are used: colors, patterns and gradients ([None] and [Registration] are skipped)
- For Library, choose the kind: Patterns, Patterns (Classic), Gradients or Gradients (Classic) (dimmed for Document)

#### Libraries (right column)

- The libraries of the chosen kind are listed with checkboxes, in the same order as Illustrator's Open Swatch Library menu (the classic kinds have dividers between groups such as Basic, Nature and Decorative)
- One block of samples is created for each checked library, stacked vertically (blocks are spaced twice the Spacing)
- Option-click a library to check only that one (Option-click again to check all)
- Select All / Deselect All toggle the checkboxes of the list shown. The count shows "checked / total"

#### Rectangle

- Set the size with Width and Height (the link icon links them)
- Set the gap between the rectangles with Spacing
- Ten rectangles per row, then wrap

#### Stroke

- Set Weight (0 for no stroke)
- Click the Color swatch to choose the stroke color with Illustrator's Color Picker (black by default)

#### Add Text

- Library Name (Document Name) adds the library name (the document name for document swatches) above the top-left rectangle of each block (14 pt)
- Swatch Name adds the swatch name above each rectangle (5 pt)
- The font is Hiragino Sans W3 (HiraginoSans-W3), or the default font if it is not installed. Kerning is set to Metrics with proportional metrics on. In Japanese versions the mojikumi set is Tsume-gumi

#### General

- Preview (Rectangles Only) draws the layout of the rectangles in the document as you change the settings (on by default). The preview does not fill them; it draws only the rectangles, strokes and text (a 0.5 pt stroke when the weight is 0)
- Number fields can be changed with the up/down buttons and the arrow keys (Shift snaps to 10s, Option steps by 0.1). Width, height and spacing use the ruler units; the weight uses the stroke units
- The samples are centered on the active artboard and left selected
- The settings are kept until Illustrator quits (library checkboxes per kind)

### How to Use

1. Open the document for the samples (for Document > Selected Swatches, select swatches in the Swatches panel first).
2. Run the script.
3. Choose the source and kind, check libraries in the list on the right, set the rectangle, stroke and text options, and click OK.

### Notes

- Library swatches are read by opening the library .ai and closing it without saving, so a window opens briefly (Illustrator cannot suspend screen redraws or open a document hidden). The swatches read are added to the Swatches panel of the current document.
- When the document already has a swatch of the same name and kind, that swatch is used instead of importing (it is not replaced even if it differs from the library).
- Preview reads only the swatch names of the libraries and imports no swatches. The names are kept until Illustrator quits, so a library is opened only the first time it is shown after Illustrator starts (it is read again when the file changes). The swatches are imported and applied only on OK. Cancel removes everything drawn for the preview.
- When the current layer (or its parent) is locked or hidden, the samples are drawn on a new layer.
- The names of the classic folders in English versions of Illustrator (Patterns (Classic) and so on) have not been verified.
- The defaults can be changed at the top of the script: `DEFAULT_SQUARE_WIDTH`, `DEFAULT_SQUARE_HEIGHT`, `DEFAULT_SIZE_LINKED`, `DEFAULT_SQUARE_MARGIN`, `DEFAULT_STROKE_WIDTH`, `DEFAULT_ADD_TITLE`, `DEFAULT_SHOW_SWATCH_NAMES`, `DEFAULT_PREVIEW`; the count per row and the text with `SQUARES_PER_ROW`, `TITLE_FONT_SIZE`, `SWATCH_NAME_FONT_SIZE`, `TEXT_FONT_NAME`, `APPLY_TEXT_MOJIKUMI` (Tsume-gumi in Japanese versions).

### Article

[Introduction article (Japanese)｜DTP Transit 別館](https://note.com/dtp_tranist/n/nbf88198f1581)

### Update History

- v1.1.2 (2026-10-27) Fixed a "Target layer cannot be modified" error that could occur on OK (a swatch already in the document with the same name is reused; the others are imported through a temporary layer). The swatch names read for the preview are kept until Illustrator quits, so libraries are not opened again on later runs
- v1.1.1 (2026-10-27) Option-click in the library list checks only that library (again to check all). The checkboxes of kinds not shown are now remembered too. The text uses Metrics kerning with proportional metrics on, and the Tsume-gumi mojikumi set in Japanese versions. Fixed an MRAP error that could occur on OK. Renamed Margin to Spacing and Library Name to Library Name (Document Name)
- v1.1.0 (2026-10-27) Added a dialog box. Several of Illustrator's bundled swatch libraries (patterns and gradients, each with its classic set) can be checked in the list in the right column (one block per library, stacked vertically). Document swatches can be limited to Selected Swatches or All Swatches and are used regardless of kind. Added Width, Height (linked) and Spacing for the rectangles, Weight and Color for the stroke, Add Text (Library Name (Document Name), Swatch Name) and Preview (Rectangles Only). Fixed the position when the rulers use artboard coordinates
- v1.0.0 (2026-10-27) Initial release
