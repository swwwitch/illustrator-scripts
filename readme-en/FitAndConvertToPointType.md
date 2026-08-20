# FitAndConvertToPointType


[![Direct](https://img.shields.io/badge/Direct%20Link-FitAndConvertToPointType.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/FitAndConvertToPointType.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### Overview

Converts the selected area type to point text. Frames whose text is overset get auto-sizing turned on first, so the overflow is cleared and no characters are lost in the conversion. Turn on "Remove forced line breaks" in the dialog to strip the breaks that remain after the conversion and join the lines.

### Features

- Picks just the area type out of the selection (locked and hidden frames, and frames on locked or hidden layers, are skipped)
- Turns auto-sizing on for the overset frames only, clearing the overflow
- Converts the frames to point text
- "Remove forced line breaks" strips the forced line breaks (soft returns) left after the conversion and joins the lines (paragraph returns are kept)
- Re-selects the resulting point text
- Reports the count when some frames could not be converted

### Flow

1. Check the document, the selection, and whether the conversion API is available
2. Show the dialog ("Remove forced line breaks")
3. Load a temporary action set for auto-sizing
4. Select each overset frame and run auto-sizing on it
5. Convert to point text (and strip the forced line breaks afterwards when the option is on)
6. Unload the action set and select the resulting point text

### Notes

- Auto-sizing cannot be set from the DOM, so it is applied through a temporary action set. The set is named after the script (`FitAndConvertToPointType_AutoSize`) so it never collides with your own actions, and it is unloaded on exit whether the run succeeds or not.
- Frames that are not overset are left alone: auto-sizing them would shift the text when it is centred or bottom-aligned.
- "Remove forced line breaks" removes the soft returns only (U+0003 / U+000A / U+2028); paragraph returns (U+000D) are kept. The characters are removed one at a time rather than by rewriting `contents`, so the per-character formatting survives.
- The option starts off. Set `DEFAULT_REMOVE_LINE_BREAKS` to `true` at the top of the script to have it on by default.
- Threaded text and anything else Illustrator cannot convert to point text is skipped, and the script reports how many frames were converted and how many were not.
- On an Illustrator without the point-text conversion API, the script says so and changes nothing.
- The conversion API does not return the resulting object, so the targets are tagged with a temporary marker name and collected afterwards. Their original names are restored once the run is done.

### Change Log

- v1.0.0 (20260820) : Initial release (with the "Remove forced line breaks" option in the dialog)
