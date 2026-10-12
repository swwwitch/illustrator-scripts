# Make a newspaper-style headline

[![Direct](https://img.shields.io/badge/Direct%20Link-ShimbunTitleMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ShimbunTitleMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

This script surrounds the selected headline (text or group) with rules and newspaper-style dummy columns, so the result looks like a clipping from a newspaper page.

The dummy can be set vertically or horizontally, in Japanese or Western text, and English newspaper mode switches everything to an English-newspaper look in one go. The dummy text is blurred, the whole piece is tilted with the Transform effect and masked by a rectangle that leaves no corner missing after the rotation. Every change is reflected in the preview, so you can check the result without closing the dialog.

<img alt="The Newspaper Headline Maker dialog" src="../png/ss-528-1008-144-20260924-012908.png" width="40%" />

### Main features

- Headline rules: Double, Single (both), Single (top only) or None, with weight and gap. Normally vertical rules on both sides of the headline; horizontal Western draws them above and below
- Margins: vertical and horizontal distance from the headline to the rules (linkable with the link icon)
- Direction: vertical (the rule height split into tiers) or horizontal (justified columns across the headline width)
- Japanese / Western: text and font of the dummy (Hiragino Mincho W3 / Times New Roman). The text can be replaced with the Text... button, separately for each
- English newspaper mode: locks Horizontal and Western, with one rule above the headline, vertical rules between columns, and side columns that run top to bottom as single tall columns
- Tiers, Gutter (a length), Extra tiers (added above and below), and tier rules with their own weight
- Type size calculated from the characters per line and tiers. In vertical direction, the side spacing is nudged so the lines of the tiers above and below align with the side tiers
- A paper-colored rectangle (Y10 K25) behind the dummy text
- Gaussian Blur on the dummy text only; the headline and rules stay sharp
- Everything grouped and rotated around its center with the Transform effect
- A mask that leaves no corner missing after the rotation, its height set by how far it reaches beyond the rules (the blurred fringe stays outside); can be turned off
- Initial choices follow the selected headline (all Western: Western and horizontal; horizontal Japanese: 1 tier; vertical: Western unavailable)
- A preview that shows the result without closing the dialog (built on a copy; the original is left untouched). Fit to Window zooms so the preview fits
- Japanese and English UI

### Usage

1. Select one headline text or group.
2. Run `ShimbunTitleMaker.jsx`.
3. Set the Margins, Rules and Mask and Rotation panels on the left and the Typesetting and Dummy Text panels on the right, and check the preview.
4. Click OK.

### Options

| Panel | Item | Default | Description |
| --- | --- | --- | --- |
| Margins | Vertical | 5 pt | How far the rules and the dummy height extend beyond the top and bottom of the headline. Horizontal Western: distance between the headline and the rules |
| Margins | Horizontal | 8 pt | Distance between the headline and the inner rule (in vertical direction it may widen slightly so the dummy lines align). Horizontal Western: how far the rules extend past the headline |
| Margins | Link | Off | When the link icon is on, the vertical and horizontal margins stay equal |
| Rules | Style | Double | Double, Single (both), Single (top only) or None. Single (top only) draws one rule above the headline and is available for horizontal Western only (switching away falls back to Single (both)). With None, the dummy text is placed where it would be with rules |
| Rules | Weight | 0.6 pt | Stroke weight of the headline rules, in the stroke unit from Preferences. Disabled with None |
| Rules | Gap | 5 pt | Distance between the inner and outer rule (Double only); 0 draws a single rule |
| Mask and Rotation | Rotation | 3° | Rotation applied to the whole piece with the Transform effect; 0 adds no effect |
| Mask and Rotation | Mask | On | Masks the whole piece with a rectangle that leaves no corner missing after the rotation; when off, no clip group is made. Turns off automatically when Extra tiers is set to 0 |
| Mask and Rotation | Horizontal | 30 mm | Width of the dummy text on each side (trimmed to whole lines). Dimmed while Mask is off, but still used as the dummy width |
| Mask and Rotation | Vertical | 10 mm | Limits the mask to this distance beyond the top and bottom of the rules (capped at the largest size that leaves no corner missing). Dimmed while Mask is off |
| Typesetting | Direction | Vertical | Vertical (the rule height split into tiers) or horizontal (columns across the headline width). Starts horizontal when the selected text is all Western |
| Typesetting | English newspaper mode | Off | While on, locks Horizontal and Western (Direction and Japanese / Western are dimmed) and sets the rules to Single (top only). The side columns run through the bands above and below as single tall columns. Turning it off restores the previous direction and script. Not available for vertical headlines |
| Typesetting | Chars per line | 11 chars | Characters in one line of a tier or column (disabled for Western horizontal) |
| Typesetting | Tiers | 3 | Vertical: number of tiers the rule height is split into. Horizontal: how many columns the headline spans (sets the column width and type size). Kept separately for each. Starts at 1 when the selected text is horizontal Japanese; 0 adds no dummy text |
| Typesetting | Gutter | 3 mm | Space between tiers. Tier rules run through its middle |
| Typesetting | Extra tiers | 1 tier | Full-width tiers added above and below each (in horizontal direction, bands of columns as tall as the rules) |
| Typesetting | Tier rules | On | Draws rules between the tiers. Horizontal Japanese: between the headline and the bands above and below. Horizontal Western: vertical rules between columns |
| Typesetting | Weight | 0.3 pt | Stroke weight of the tier rules, in the stroke unit from Preferences. Dimmed while Tier rules is off |
| Typesetting | Blur | On, 8 px | Check to apply a Gaussian Blur to the dummy text. The value is its radius; 0 also applies none. Option (Alt)+B toggles it |
| Dummy Text | Japanese / Western | Japanese | Text and font of the dummy (Hiragino Mincho W3 / Times New Roman). Starts as Western when the selected text is all Western. Western is unavailable for vertical headlines |
| Dummy Text | Text... | Lorem ipsum… | Opens a dialog to set the text poured into the dummy, for Japanese and Western; it is repeated as needed. Empty uses the default text; line breaks are removed |
| Button row | Fit to Window | On, 65% | Zooms so the preview fills the given share of the window (10–100%). Cancel restores the original view |

Lengths are shown in the ruler unit and weights in the stroke unit, inside the field (such as "5 mm"; values typed in another unit are converted, and a bare number gets the field unit). The numeric fields step with the stepper buttons to their left or the Up/Down arrow keys (to the next whole number rather than by 1, e.g. 1.5 → 2; Shift: to the next multiple of ten; Option: ±0.1; Chars per line, Tiers and Extra tiers take whole numbers only). Only Rotation accepts negative values.

### Dummy text dimensions

In vertical direction, the dummy is set as follows.

| Item | Value |
| --- | --- |
| Character advance | (Rule height − (tiers − 1) × gutter) ÷ (tiers × characters per line) |
| Type size | Character advance ÷ vertical scale (85% for Japanese, 100% for Western) |
| Leading | Type size × 1.4 |
| Between tiers | Gutter, with a tier rule (Weight) in the middle |
| Space next to the rules | One character |
| Font | Hiragino Mincho W3 for Japanese, Times New Roman for Western (the default font when missing). Western gets extra text for its narrower glyphs |

In horizontal direction, the dummy is set as follows.

| Item | Value |
| --- | --- |
| Column width | Characters per line × average character width (Japanese: Chars per line × 1 em; Western: about 32 × 0.5 em) |
| Type size | The size at which the span between the outer rules equals tiers × column width + (tiers − 1) × gutter |
| Gutter and space next to the rules | Gutter |
| Leading | Type size × 1.4 (Japanese) / 1.2 (Western) |
| Bands above and below | As tall as the rules. Japanese gets a rule (Weight) between them and the headline |
| Between columns (Western) | A vertical rule (Weight) in the middle |
| Font | Hiragino Mincho W3 for Japanese (85% vertical scale), Times New Roman for Western (then Times, Georgia, or the default font) |

The background is a Y10 K25 rectangle in both cases. The rule and background colors, the font, the vertical scale, the leading and so on can be changed under "User Settings" at the top of the script. In RGB documents, the CMYK values are converted to RGB with a simple formula.

### Resulting objects

From front to back, everything is wrapped in one clip group (with Mask off, only the group in item 2):

1. The mask rectangle
2. The group rotated with the Transform effect
   - The headline
   - The rules (`Side rules`; not made with None)
   - The dummy text (`Newspaper dummy`, blurred)
   - The background rectangle

### Notes

- Works with Illustrator 2024–2026.
- Select exactly one text object or one group before running. Both vertical and horizontal headlines work.
- Western and English newspaper mode are unavailable for vertical headlines.
- The dummy text repeats the opening of "I Am a Cat" (Japanese) or Lorem ipsum (Western). For Japanese, kinsoku is set to None so every line keeps the set length (a period may start a line).
- Each tier or column is filled with more text than fits; the overflow stays hidden.
- The blur and the rotation are live effects, so they can be adjusted later in the Appearance panel.
- While previewing, the selected headline is hidden and a copy is shown. Cancel restores the original state and view.

### Article

https://note.com/dtp_tranist/n/ndb9bee6b7a2e

### Change log

- v1.5.0 (2026-10-12): Reorganized the dialog into the Margins, Rules, Mask and Rotation, Typesetting and Dummy Text panels. Added Direction (vertical / horizontal) and English newspaper mode, so Japanese / Western and the direction are chosen separately (horizontal Western draws the headline rules above and below and the tier rules as vertical rules). Added Single (top only) to the rule styles, Weight for the rules and the tier rules, Gutter, Vertical for the mask, linked margins, Text... (replace the dummy text), Tier rules, Fit to Window, and Option (Alt)+B to toggle Blur. Merged Tiers and Columns into one field. Initial choices now follow the selected headline. Length and weight fields hold their unit, and label widths are measured and aligned. Reviewed the wording and cleaned up the internals
- v1.4.1 (2026-10-12): Added a checkbox to Blur so it can be turned on and off without changing the radius
- v1.4.0 (2026-10-12): Added Japanese and Western options to the newspaper dummy; Western builds horizontal columns from the number of columns the headline spans. The dialog now has two columns, with the newspaper dummy settings on the right
- v1.3.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.3.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.3.5 (2026-10-01): The button row, previously always centered, is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones. Unified the window and panel margins and spacing with the shared layout part
- v1.3.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.3.3 (2026-09-29): Dialog opacity changed to 98%
- v1.3.2 (2026-09-28): The button row is now built with the shared part
- v1.3.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.0 (2026-09-24): Added Mask to Options (off: no clip group; turns off automatically when Extra tiers is set to 0)
- v1.1.1 (2026-09-24): Added periods to the dummy text (kinsoku set to None); widened the label column so long labels are not clipped
- v1.1.0 (2026-09-24): Added the dummy text (tiers, tiers above and below, blur, background), rotation, mask and preview; the margins are now two items, vertical and horizontal
- v1.0.0 (2026-09-23): Initial release
