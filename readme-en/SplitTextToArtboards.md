# Split text by paragraph and place each on its own artboard

[![Direct](https://img.shields.io/badge/Direct%20Link-SplitTextToArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SplitTextToArtboards.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitTextToArtboards.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sometimes you want each paragraph of a multi-line text on an artboard of its own: slide headings, name tags, or a set of title ideas to compare. By hand, that means duplicating the text for every paragraph, deleting the other lines, scaling it to the artboard, and centering it, over and over.

This script splits the selected text into one text object per paragraph. It places them on the artboards in order, starting from the one you choose, and fits each to the artboard width, centered. When there are not enough artboards, it adds more that follow the existing layout.

### Features

- Splits the selected text into one text object per paragraph, keeping the formatting
- Places the paragraphs one per artboard, in order, starting from the artboard you choose
- Scales each text object to a percentage of the artboard width (90% by default), keeping its aspect ratio, and centers it on the artboard
- Adds artboards automatically when there are not enough. The column count and gap are read from the existing layout and used as defaults
- Skips empty paragraphs (paragraphs of only whitespace and line breaks)
- Keeps or deletes the original text, as you choose

### How to use

1. Select a text object that contains several paragraphs. If you have some characters selected, the text object that contains them is used
2. Run the script, check the settings in the dialog, and click OK

When the script finishes, the new text objects are selected.

The bottom of the dialog shows the number of paragraphs to place and the current number of artboards.

### Options

| Option | Description |
|---|---|
| Start from artboard | Number of the artboard for the first paragraph (counted from 1) |
| Artboard width | Target width as a percentage of the artboard width. 100 fills the full width |
| Columns when adding | Number of columns used when adding artboards. Prefilled from the existing layout when it can be read (6 otherwise) |
| Artboard gap | Gap between added artboards, in pt. Prefilled from the existing layout when it can be read (20 pt otherwise) |
| Keep the original text | When on, the source text is kept instead of deleted (off by default) |

In the number fields, Up/Down changes the value by 1, and Shift+Up/Down by 10.

These can also be changed in the user settings at the top of the script:

| Setting | Description |
|---|---|
| `USE_PREVIEW_BOUNDS` | `true` measures by preview bounds (incl. strokes & effects), `false` by geometric bounds (default `true`) |
| `CENTER_VERTICALLY` | Set to `false` to keep the vertical position instead of centering vertically (default `true`) |
| `SKIP_EMPTY_PARAGRAPHS` | Set to `false` to use an artboard for each empty paragraph as well (default `true`) |

### Notes

- Only the first text object found in the selection is processed
- Added artboards are the same size as the last artboard. They continue to the right of it and wrap at the column count
- Artboards cannot be added outside the canvas. If the script reaches the edge of the canvas and cannot place every paragraph, it stops and tells you how many paragraphs it placed
- The target size is based on the width only. On a tall artboard, the text can overflow the artboard height

### Version history

- v1.0.0 (2026-09-01): Initial version
