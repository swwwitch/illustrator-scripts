# Build a blur, a loupe or a detail callout from a masked copy

[![Direct](https://img.shields.io/badge/Direct%20Link-MaskSpotlight.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/MaskSpotlight.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MaskSpotlight.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Select an image or a group of vector objects together with the path to mask with, then run the script. It masks a copy of the artwork with that path and builds a frosted-glass panel, a loupe or a detail callout, all against a live preview.

The original artwork and path stay as they are, and the only additions are one copy and the clipping group — the single exception being Outside the mask, which blurs the original artwork itself.

### Features

- Blur area set to None, Inside the mask or Outside the mask (radius 0 to 1000 px)
- Six presets: blurred foreground, blurred background, mask with drop shadow, loupe, detail callout, and detail callout with a blurred background
- Saves the current settings and offers them as Saved settings the next time
- A stroke on the clipping group, and a drop shadow (opacity, X, Y, blur, darkness)
- Loupe: a scale plus a move measured against the mask size, applied as a real duplicate or as a Transform effect
- A zoom connector that links the artwork and the magnified copy with lines
- A preview that hides the original artwork and stands in for it
- Japanese and English UI

### Usage

1. Select one image (placed or embedded) or group, and one path or compound path to mask with.
2. Run `MaskSpotlight.jsx`.
3. Pick a preset or set the values yourself (the dialog opens on Blurred background).
4. Check the preview and click OK.

### Settings

| Setting | What it does | Default |
| --- | --- | --- |
| Preset | Loads a set of settings at once; changing anything switches it to Custom | Blurred background |
| Blur area | None, Inside the mask or Outside the mask | Outside the mask |
| Radius | Radius of the Gaussian blur (0 to 1000 px) | 15 px |
| Add stroke | Adds a stroke to the clipping group and excludes it against itself | Off |
| Zoom connector | Applies the graphic style that links the artwork and the magnified copy | Off |
| Add drop shadow | Applies a drop shadow (Multiply) to the clipping group | Off |
| Opacity | Opacity of the shadow (0 to 100%) | 50% |
| X / Y (shadow) | How far the shadow shifts (±1000 pt; positive moves right and down) | 5 pt / 5 pt |
| Blur | Blur of the shadow (0 to 1000 pt) | 8 pt |
| Darkness | Darkness of the shadow (0 to 100%); it uses darkness rather than a color | 100% |
| Add loupe | Scales the mask up and lays it over the artwork like a magnifier | Off |
| Duplicate | Keeps the original glass panel and adds the loupe beside it | Off |
| Scale as an effect | Applies the scaling and the move as a Transform effect, leaving the object untouched | Off |
| Scale | How far the loupe is scaled up (100 to 1000%) | 120% |
| X / Y (loupe) | Moves the loupe as a share of the mask width and height (±200%; positive Y moves down) | 0 / 0 |
| Preview | Shows the result on the document while you adjust the settings | On |

### Notes

- The selection must be exactly one image or group plus one path or compound path; anything else stops the script.
- The loupe is unavailable while the blur area is Inside the mask, since there would be nothing sharp left to magnify — the panel dims.
- The zoom connector is available only where the loupe is duplicated, that is with the Detail callout presets. Every other setting dims it.
- The connector uses the graphic style `zoomineffect` from `graphiclibraryzoomin.ai`. Without that file beside the script, the option is not shown at all.
- A graphic style replaces the appearance of whatever it lands on, so the connector groups the result first and applies the style to that group, leaving the stroke, the drop shadow and the Transform effect intact.
- When the artwork is a vector group, the clipping group also gets the Divide option of the Pathfinder effect. An image has no paths to divide, so it is left out.
- Add stroke runs menu commands, which are not available in every environment.
- The saved settings live in `~/Library/Application Support/MaskSpotlight/settings.txt` (`%APPDATA%\MaskSpotlight\settings.txt` on Windows).

### Article

[DTP Transit 別館｜note](https://note.com/dtp_tranist/n/nfc777dda965d)

### Changelog

- v1.0.0 (2026-09-20): First release
