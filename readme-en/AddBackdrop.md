# Lay a shape behind text or objects

[![Direct](https://img.shields.io/badge/Direct%20Link-AddBackdrop.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/AddBackdrop.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Lays a circle, superellipse or rectangle behind the selected text or objects.

The shape is sized from the visual bounds of the text as if it were outlined. When an existing backdrop is selected too, the script reads its shape and replaces it.

<img alt="The Add Backdrop dialog" src="../png/ss-990-1102-144-20260923-190826.png" width="50%" />

### Features

- Three shapes: circle, superellipse and rectangle (also switchable with the E / S / R keys)
- Rectangles take margins, a square option, rounded corners and a pill shape
- Fill or stroke, color (text color / black / white / CMYK) and opacity
- Groups the backdrop with the text, or knocks the text out with Exclude
- Live preview while the dialog is open; OK commits everything as a single undo step
- Settings are remembered until Illustrator quits

### Usage

1. Select the text or objects to put a backdrop behind.
2. Run the script.
3. Set the shape and options, check the preview, and click OK.

**How the target is chosen**

- If the selection contains text, the first text frame (including one inside a group) is the target.
- Otherwise, the bounds of the whole selection are used. Single Character and Text Color are unavailable in that case.

**Replacing an existing backdrop**

With either of these selections, the script treats the shape as the existing backdrop and reflects its type in the shape buttons. On OK, the old shape is removed and replaced.

- A single group containing the text and a shape
- The text (or a group) selected together with a shape

### Options

| Option | What it does |
| --- | --- |
| Scale: Size | Size of the shape relative to the base bounds, in percent. Defaults to 90%. Fixed at 100% for rectangles (except squares) |
| Scale: Single Character | Draws a circle whose diameter is 1.5 times the font size. Turned on automatically for text with a single non-space character |
| Margin: Vertical / Horizontal | Space added above/below and left/right of the base bounds. Defaults to a quarter of the short side on first use |
| Margin: Link | Uses the vertical value for the horizontal margin too |
| Margin: Square | Makes a square whose side equals the circle's diameter |
| Round | Corner radius, applied with the Round Corners live effect. Defaults to a fifth of the short side on first use |
| Pill shape | Uses half the height as the radius so both ends become semicircles |
| Group with Text | Groups the backdrop with the target |
| Exclude | Knocks the text out with Live Pathfinder Exclude. Also turns on grouping and Text Color. Unavailable for strokes |
| Offset: X / Y | Nudges the shape, in ruler units |
| Fill & Stroke | Fill or stroke the shape; set the weight for a stroke |
| Color | Text Color / Black / White / CMYK |
| Opacity | When on, sets the opacity of the shape |

Margins, Square, Round and Pill shape apply to rectangles only.

The Show Transparency Grid button toggles the transparency grid, handy for checking a white backdrop.

**Keys**

| Key | Action |
| --- | --- |
| E / S / R | Switch to circle / superellipse / rectangle |
| Up / Down | ±1 |
| Shift + Up / Down | ±10 (snaps to multiples of 10) |
| Option + Up / Down | ±0.1 |

### Notes

- On commit, a circle is converted to a live shape with Convert to Shape, a rectangle becomes a compound shape via Live Pathfinder Add, and a superellipse stays an 8-point path.
- The preview relies on Illustrator's undo history.

### Article

https://note.com/dtp_tranist/n/na8af4a7016ad

### Update History

- v1.6.4 (2026-09-23): Revised UI wording (title, panel names, label colons, tooltips); split the button row and added a Show Transparency Grid button
- v1.6.3 (2026-09-23): Fixed typed stroke weight and CMYK values not being clamped, the previous corner radius not being restored when rounding was on, and the message shown with no document open. Internal cleanup
- v1.6.1 (2026-03-26)
