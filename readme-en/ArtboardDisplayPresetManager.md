# ArtboardDisplayPresetManager

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardDisplayPresetManager.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/ArtboardDisplayPresetManager.jsx)

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
| Width / Height | Shown in the current ruler unit. Edit a value and commit to resize the artboard (anchored at its top-left corner). |
| Optimize to Pixel Grid | Rounds the artboard's XYWH to integers. |
| Reload | Re-reads the current artboard info. |

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
