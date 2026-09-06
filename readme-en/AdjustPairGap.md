# Set the gap and position from a side you fix

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustPairGap.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AdjustPairGap.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustPairGap.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Sets the gap and the position of the selected objects to a value you specify. The side marked as fixed stays put and only the rest move. The preview on the artboard updates as you change values and options, and OK keeps exactly what you see.

### Main Features

- Three modes: even spacing inside a group, nearest-neighbour pairs, and margins from an artboard edge
- Key Object is a cross of top/left/right/bottom. Left or right adjusts horizontally, top or bottom vertically
- The gap accepts negative values, so objects can overlap
- A single Position panel that switches between horizontal and vertical with the Key Object side, relabelling its radios left/right or top/bottom
- Alignment perpendicular to the gap (none / left (top) / center / right (bottom)), plus an extra offset from there
- Text alignment (auto / left / center / right / justify) as icon buttons
- Preview Bounds switches between visible bounds (stroke and effects included) and geometric bounds
- Arrow keys step the numeric fields (Shift by 10, Option by 0.1)
- Keyboard shortcuts for alignment (L / C / R for the horizontal direction, T / M / B for the vertical one)
- Values follow the current ruler unit
- Remembers the dialog state and restores it next time
- Japanese and English UI

### Usage

1. Select two or more objects.
2. Run the script.
3. Choose the mode, the key object and the gap, check the preview, and click OK.

### Modes

| Mode | What it does |
| --- | --- |
| Group | Lays out the contents of each selected group so every adjacent gap is equal (3+ objects supported) |
| Auto Pair Detection | Pairs the selected objects by nearest neighbour and sets the gap of each pair |
| Artboard | Sets each selected object's gap (margin) to the artboard edge chosen in Key Object (top/left/right/bottom) |

Group is preselected when every selected object is a group, otherwise Auto Pair Detection (a remembered setting wins over both).

### Options

**Key Object**

Radios arranged as a cross of top/left/right/bottom. The object on the chosen side stays put and the rest move relative to it.

| Chosen side | Gap direction | Position panel |
| --- | --- | --- |
| Left / Right | Horizontally | Vertical |
| Top / Bottom | Vertically | Horizontal |

**Offset**

The panel title carries the current ruler unit, e.g. `Offset (mm)`. The Position panel does the same.

| Item | What it does |
| --- | --- |
| Gap | The gap between objects; a negative value overlaps them. Starts at the selection's current average gap |
| Preview Bounds | On: measure by visible bounds (incl. stroke/effects). Off: geometric bounds |

**Position**

There is one panel, and it switches to whichever direction is perpendicular to the Key Object side. Which orientation is live shows in the radio labels — left/right for horizontal, top/bottom for vertical. Its title carries the current ruler unit. The horizontal and vertical settings are remembered separately, so switching the key side away and back brings your values back.

| Item | What it does |
| --- | --- |
| Align | Aligns the moving side to the key object's left (top) / center / right (bottom). "None" leaves it alone |
| Position | Nudges the moving side further after alignment (positive = right for horizontal, down for vertical). Held at 0 while Center is selected |

**Text alignment**

Five icon buttons; the active one is drawn inverted.

| Button | What it does |
| --- | --- |
| A (Auto) | Area text is justified. Point text follows the Position panel's alignment when stacked vertically, or the Key Object side (left/right) when laid out horizontally |
| Left / Center / Right / Justify | Applies the same justification to every selected text (Justify aligns the last line left) |

### Notes

- At least two objects must be selected.
- Group mode skips objects that are not groups, and groups with fewer than two children.
- In Auto Pair Detection mode, an odd selection leaves one object without a partner, and it is skipped.
- Pairing always uses the centers of the geometric bounds. Preview Bounds only changes how the gap is measured, never which objects are paired.
- Clip groups are measured against their clipping path, so artwork hidden by the mask is not counted.
- Changing the justification keeps the text's visual position (the frame is moved back so point text does not shift around its anchor).
- The live preview is the result. Cancel reverts both the positions and the text justification.
- The alignment shortcuts (L / C / R, T / M / B) do nothing while a modifier key is held, so Cmd+C and friends keep working.
- All settings are remembered for the Illustrator session. Mode and gap depend on the selection, so they are re-derived from it after an Illustrator restart.

### note

https://note.com/dtp_tranist/n/nc8fab19d8164

### Update History

- v1.0.0 (20260608): Initial release (auto pair detection, preview bounds, arrow-key stepping)
- v1.1.0 (20260609): Added group mode
- v1.2.1 (20260610): Added the vertical axis and the alignment (moved side) panel
- v1.2.2 (20260611): Added settings persistence and clip-group bounds handling
- v1.3.0 (20260628): Added artboard mode, text alignment, position offsets and alignment shortcuts
- v1.3.1 (20260629): Reorganized the dialog naming and panels
- v1.3.2 (20260906): Merged the Horizontal/Vertical panels into a single Position panel, renamed the old Position panel to Offset and moved the unit into its title, turned text alignment into icon buttons, and fixed modified keystrokes being swallowed and text shifting on justification changes
