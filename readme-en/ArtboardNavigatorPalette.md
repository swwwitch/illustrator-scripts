# Move between artboards from a palette

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardNavigatorPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ArtboardNavigatorPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardNavigatorPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A palette for moving smoothly between artboards, zooming as it goes.

### Features

- `|<` jumps to the first artboard (disabled while you are already there)
- `<` goes to the previous artboard, wrapping at the ends (Ctrl+Shift+Opt+Left)
- `■■` fits every artboard on screen (Ctrl+Shift+Opt+Up)
- `>` goes to the next artboard, wrapping at the ends (Ctrl+Shift+Opt+Right)
- `>|` jumps to the last artboard (disabled while you are already there)
- Buttons brighten slightly while pressed, as click feedback

### Usage

1. Run the script to open the palette.
2. Use the navigation buttons, or the keyboard shortcuts.

### Options

| Option | Effect |
| --- | --- |
| Animation | When off, the view switches instantly with no interpolation (the options below are dimmed) |
| Speed | How fast the move animates (1–10; further right is faster, with fewer steps) |
| Ease out | When on, the move decelerates toward the end (easeOut); when off, it moves at a constant speed |
| Prezi-like mode | On previous/next moves, zooms out once midway before closing in |
| Overview strength | How far Prezi-like mode zooms out (0–1) |
| Show artboard name | Draws a label on the destination artboard (when off, nothing is drawn and no layer is created) |

These settings and the palette position carry over to the next session.

### Artboard list

- A checkbox shows or hides the list. When hidden, the palette shrinks instead of leaving empty space.
- The list shows number and name; click a row to move to that artboard.
- The palette is resizable, and the list grows with it.

### Artboard label (when "Show artboard name" is on)

- Right after moving to an artboard, "number: artboard name" appears at its top left.
- White HiraginoSans-W6 text on a black rectangle.
- Text size and the background/text opacity can be adjusted with the variables at the top of the script (`LABEL_FONT_RATIO` / `LABEL_BACKGROUND_OPACITY` / `LABEL_TEXT_OPACITY`).
- The label is drawn on a dedicated layer, "ArtboardNavigator" (locked, non-printing), and redrawn on every move.
- The layer is removed when "Show artboard name" is turned off, when fitting all artboards, and when the palette is closed.

### Notes

- The palette is a resident script: after editing the code, close the palette before running it again, otherwise the old code keeps running.
- The target zoom is calculated from `view.bounds`, so the animation starts without a lag.
- The actual view changes and animation run in the main engine through BridgeTalk.
- The list selection is settled in the palette when you click, without waiting for BridgeTalk's `onResult`.

### Acknowledgments

Yuki Furushima contributed many ideas and much of the code, including how the BridgeTalk worker is installed and called and how the artboard labels are drawn.

https://note.com/yukifurushima/n/n9f2078dc156f

### Update History

- v1.2.14 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.2.13 (2026-10-03) Moved to the shared persistent engine `SwwwitchPalettes` so CloseAllPalettes can close it
- v1.2.12 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.2.11 (2026-09-29) Esc now closes the palette (also while typing)
- v1.2.10 (2026-09-29) Code cleanup. Field labels now end with a colon, the checkbox is renamed "Show artboard list", and the fit-all tooltip reads "Fit all artboards in the window"
- v1.2.9 (2026-09-29) Fixed a bug where the palette opened but the buttons and list no longer moved the view
- v1.2.8 (2026-09-28) Settings are now saved through the shared part (stored in `~/Library/Application Support/illustrator-scripts/ArtboardNavigatorPalette.json`; the old settings are carried over once)
- v1.2.7 (2026-09-26) Renamed the file from `ArtboardNavigator.jsx` to `ArtboardNavigatorPalette.jsx`.
- v1.2.5
