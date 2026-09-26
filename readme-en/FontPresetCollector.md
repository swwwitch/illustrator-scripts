# List the paragraph formats in a document and change or replace them in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-FontPresetCollector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontPresetCollector.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetCollector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Surveys every paragraph in the active document for its format (font, font size, leading, tracking, tsume, kerning, justification and paragraph spacing), and lists paragraphs that share a format together.

Edit the values of a format in the list, or replace it with another format, and every paragraph in that format is updated at once. It works much like redefining a paragraph style, without using styles, so you can tidy up scattered formatting.

### Key Features

- Surveys all text in the document paragraph by paragraph and lists each format (most-used by character count first)
- Each row reads "Font name  Font size / Leading  (paragraph count)"
- Opens with the format of the text selected at launch already chosen in the list
- Choose one format to edit its values under **Edit Format** on the right; **Update** applies them to that format's paragraphs right away
- Choose formats in the list (one or more), pick a format in the popup, and **Replace** them with it in one go
- After every Update or Replace the document is surveyed again: paragraphs that now share a format merge into one row and the paragraph count is recounted
- **OK** keeps the changes; **Cancel** puts everything back as it was when the dialog opened
- Number fields step with the ↑↓ keys (Shift for ±10, Option for ±0.1)
- Font size, leading and paragraph spacing are shown in the type unit set in Preferences (pt, Q, mm and so on)

### How to Use

1. Open the document you want to tidy up (select some text first and its format is chosen when the dialog opens).
2. Run the script. The formats are listed on the left.
3. To change values, choose one format, edit the fields on the right and click **Update**.
4. To match other formats, choose them in the list (Shift/Command-click for several), pick the target format in the popup below and click **Replace**.
5. Changes show on the canvas right away. Click **OK** to keep them or **Cancel** to discard them.

### Options

- **Format list** — The formats used in the document. The number in brackets is how many paragraphs use it.
  - With one format chosen, **Edit Format** on the right is available.
  - With several chosen, **Edit Format** is dimmed and only **Replace** is available.
- **Replace** — Replaces the paragraphs in the formats chosen in the list with the format picked in the popup.
  - Formats chosen in the list are dimmed in the popup and cannot be picked.
  - Every value of the picked format is copied, including whether leading is auto or fixed (and the auto-leading percentage when auto).
  - Afterwards, the picked format is chosen in the list.
- **Edit Format**
  - **Font** — Choose from the fonts used in the document.
  - **Font Size**
  - **Leading / Auto** — With **Auto** on, leading follows the font size (the auto-leading percentage is kept). Turn it off to enter a fixed value.
  - **Tracking** — In 1/1000 em.
  - **Tsume** — In percent.
  - **Kerning** — Metrics, Optical, Metrics - Roman Only or 0.
  - **Justification** — Align Left, Align Center, Align Right, Justify with Last Line Aligned Left / Center / Right, or Justify All Lines.
  - **Space Before / Space After**
  - **Update** — Applies the edited values to the paragraphs in this format right away.
  - While there are edits not yet applied with Update, the panel title reads "Edit Format (unapplied changes)".
- **OK** — Keeps the changes and closes. Edits not yet applied with Update are applied now.
- **Cancel** — Undoes every change, including those applied with Update and Replace, and closes.

### Notes

- Paragraph styles are not used. The character and paragraph attributes of each paragraph are rewritten directly.
- A format is read from the first character of the paragraph. If the formatting changes partway through a paragraph, the paragraph is still listed under the format of its first character.
- Only the settings you changed are written. Settings you left alone, and attributes the list does not cover (horizontal scale, baseline shift and so on), stay as they are. A setting that is written covers the whole paragraph, though, so values that differed partway through become the same.
- Empty paragraphs are not listed.
- When you edit a format, leading defaults to **Auto**. To keep a fixed leading, turn **Auto** off and check the value before clicking **Update**. If you leave **Auto** on but change nothing else, the leading is not rewritten.
- The auto-leading percentage cannot be edited in the fields. To change it, **Replace** with a format that has the percentage you want.
- Paragraphs that cannot be written to, such as text on a locked layer, are skipped.
- The **Font** choices are limited to the fonts used in the document. To switch to another font, place some text in that font in the document first.
- Fonts missing from your system are shown by their PostScript name and cannot be switched to (the other values are still applied).
- Values are rounded to two decimal places.

### Release Notes

- v1.0.1 (20260926) : Fixed formats not merging, and paragraph counts not being recounted, after replacing with a format that differs only in its auto-leading percentage
- v1.0.0 (20260926) : Initial release

### Script Info

- Version: v1.0.1
