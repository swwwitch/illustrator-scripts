# Manage artboard display preferences

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardDisplayPresetManagerPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/ArtboardDisplayPresetManagerPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardDisplayPresetManagerPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Artboard-related settings are scattered across several categories of the Illustrator Preferences dialog. This script gathers them into a single persistent palette.
- There is no OK button: every change is written the moment you click it, so you can fine-tune while watching the canvas.
- The palette also shows the active artboard's number, name and size, and can resize it or snap it to the pixel grid.

### Usage

1. Run the script; a persistent palette opens. You can keep working with it open.
2. Press Esc while the palette is active to close it.
3. Running the script again does not open a second palette — the existing one is brought forward.

### Current Artboard

| Item | Behavior |
| --- | --- |
| Width / Height | Shown in the current ruler unit. Edit a value and commit to resize the artboard (around the reference point chosen in the 9-axis widget on the right; top-left by default). The stepper buttons and arrow keys step to the next whole number (1.5 → 2; Shift to the next multiple of 10, Option ±0.1); a click resizes at once, the arrow keys resize on key release. |
| Optimize to Pixel Grid | Rounds the artboard's XYWH to integers. |

Zero, negative or non-numeric input is rejected and the fields revert to the current values.

The info is also refreshed whenever the palette is re-activated, so clicking the palette after switching artboards updates the display.

### Artboard Name & Border

| Item | Preference key |
| --- | --- |
| Show Artboard Name | showArtboardLabelOnCanvas |
| Highlight Color | ArtboardBBColorRed / Green / Blue |
| Stroke Width (1-4) | ArtboardBBWidth |

Nine colors are available (Light Blue, Light Red, Green, Medium Blue, Magenta, Cyan, Light Gray, Black, Yellow). If the stored color does not match a preset exactly, the **closest** one is selected.

### Options

| Item | Preference key |
| --- | --- |
| Move Locked or Hidden Objects Together | moveLockedAndHiddenArt |

"Show the 'Print Bleed' Generative AI Button" (`enablePrintBleedWidget`) was dropped. The value is written reliably and the Preferences dialog reflects it, but the canvas widget is never re-evaluated — redraw, zoom, tool switching, preview/outline toggling, artboard re-assignment and document switching all fail to apply it. The code is kept commented out in case a future Illustrator version behaves differently.

### Presets

Three radio buttons below the border panel switch all of the above at once.

| Preset | Artboard name | Color | Width | Move together |
| --- | --- | --- | --- | --- |
| Default | Shown | Black | 1 | Off |
| Emphasis | Hidden | Light Red | 3 | On |
| Light | Hidden | Light Gray | 1 | On |

When the palette opens, a preset is selected only if the current preferences match **every** value in it; otherwise no radio button is selected.

### Bottom buttons

| Button | Behavior |
| --- | --- |
| Change Canvas Color | Toggles the canvas outside the artboards between white and gray (uiCanvasIsWhite). |
| Video Ruler | Toggles the video ruler. |

### Scope

Illustrator application preferences, plus the active document's artboard (pixel-grid optimize and resize only). Existing objects are never modified.

### Notes

- Preference changes do not repaint the canvas by themselves, so a zoom-out / zoom-in pair is issued after each write to force a refresh.
- An Illustrator persistent palette loses its DOM connection while shown, so artboard reads and resizes are delegated to the main engine via BridgeTalk.
- With no document open, the artboard info shows "—". An alert appears only when you actually try to optimize or resize.

### Update History

- v1.3.5 (2026-10-10) Fixed the palette operating on another Illustrator version when several versions are running
- v1.3.4 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.3.4 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.3.3 (2026-10-03) Moved to the shared persistent engine `SwwwitchPalettes` so CloseAllPalettes can close it
- v1.3.2 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.3.1 (2026-09-28) The 3×3 reference point picker now uses the shared part.
- v1.3.0 (2026-09-27) Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten).
- v1.2.5 (2026-09-26) Renamed the file from `ArtboardDisplayPresetManager.jsx` to `ArtboardDisplayPresetManagerPalette.jsx`.
- v1.2.4 (2026-09-25) The width/height fields can now be stepped with the arrow keys (Shift ±10, Option ±0.1); the artboard is resized when the key is released. Width and height are now stacked vertically, with a 9-axis widget beside them to set the resize reference point. Removed the Reload button (info is re-read when the palette is activated).
- v1.2.3 (2026-09-25) Fixed the width/height fields showing the wrong tooltip (a preset-name hint). Added tooltips to the buttons. The unit label now reads "H" when the ruler unit is Ha.
