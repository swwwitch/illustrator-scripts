# Draw a branching Sankey diagram from the selected object

[![Direct](https://img.shields.io/badge/Direct%20Link-SankeyFlowMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/SankeyFlowMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SankeyFlowMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Uses the selected object as the left node and draws a flow diagram (Sankey diagram) that branches out from it. Band widths follow the values.

If you lay out the label texts first, they stay where they are, and the bands are fitted to their positions and font size.

### Main Features

- The node is the selected key object (or the leftmost object)
- Flows are written one "value name" per line; indent to branch
- Band widths proportional to the values or to their square roots
- Bands fitted to the label texts you laid out (the texts stay put), or the texts moved to the bands
- A rounded box is drawn behind a text node
- Circles or rectangles at junctions that enclose junction captions (separate settings for the first level and deeper levels)
- Arrow tips at the ends
- Bands can also be drawn as strokes (stroke weight = band width)
- Bands in a single color or colored by first-level branch, with per-band opacity
- Outline Stroke and Pathfinder (Merge) effects applied to the band group
- Preview; remembers the last settings
- Fit to Window zooms so the preview fits
- Japanese and English UI

### How to Use

1. Select the object to use as the node and the texts to use as labels. Making the node the key object is the most reliable.
2. Run `SankeyFlowMaker.jsx`.
3. Enter the flows under Flow, then adjust the size and junction settings. The preview redraws whenever a setting changes.
4. Click Create. The diagram is created as a "Sankey Flow" group behind the selected objects.

### Writing the Flows

Write one "value name" per line. The value may come first or last (`9280 inside sandbox` / `inside sandbox, 9280`). A line indented deeper (with spaces) branches from the shallower line above it. Lines starting with `#` are ignored.

```
9280 inside sandbox
720 auto-reviewed
  713 approved and continue
  7 denied
    4 continue via safer alternative
    3 stop and ask user
```

### How Label Texts Are Used

Selected texts (other than the node) are used as follows. Matching ignores case, spaces, thousands separators, and punctuation.

| Text | Used as |
| --- | --- |
| "value name" or "name" matches a branch | That branch's label. The text stays put; the band's center, curve end, and arrow tip are fitted to it. |
| Remaining texts without digits | Junction captions, assigned to junctions from the leftmost; the junction is placed at the text. |
| Branches with no matching text | A new label is created in the format of the selected label texts. |

### Size Panel

| Item | Description |
| --- | --- |
| Width Scale | Proportional or Square Root. Square Root keeps small values visible. |
| Node Height | Height of the rounded box behind a text node. Fixed to the shape's height for a shape node. |
| Exit Height | Band exit height (total width of the first-level bands) as a percentage of the node height. The node keeps its size. |
| Curve Length | Horizontal length of the curved part. |
| Straight Length | Length of the straight part after the curve. |
| Branch Gap | Gap between neighboring branches. |
| Min Width | Bands thinner than this are drawn at this width. |
| Font Size | Font size of new labels. Fixed to the label texts' size when they are selected. |

Where a branch has a text to fit, the text's position takes precedence over Curve Length and Straight Length.

### Options Panel

| Item | Description |
| --- | --- |
| Show Values in Labels | New labels read "9,280 inside sandbox"; off shows the name only. |
| Arrow Tips | Points the ends of bands that have no children. |
| Draw Bands as Strokes | Draws each band as a stroke along its center (stroke weight = band width, butt caps) instead of a filled shape, so the width can be changed later via the stroke weight. Arrow tips become filled triangles following the stroke. On steep curves the stroke looks thinner vertically. |
| Move Labels | On: lays out the bands by the set lengths and gaps and moves the selected label texts to fit (junction captions go to the junction centers). Off (default): the texts stay put and the bands are fitted to them. Canceling puts the texts back. |

### Junctions (Level 1) / Junctions (Level 2+) Panels

Junctions (Level 1) covers junctions at the end of first-level branches; Junctions (Level 2+) covers second-level and deeper branches.

| Item | Description |
| --- | --- |
| Circle / Rectangle / None | Shape placed at the junction. With a junction caption, it is sized to hold the text, and the text is center-justified without moving (Cancel or None restores the original justification). |
| By Size / By Margin | What width and height mean: the shape size, or the margins left/right and above/below the junction caption. |
| Width, Height | The shape size or the margins. 0 = auto (sized to hold the text) for both. |
| Link (chain icon) | Keeps the width and height equal. On by default. |

When a new label is created for a middle branch whose label does not fit in the band, it goes inside the junction shape.

### Color Panel

| Item | Description |
| --- | --- |
| Single Color / By First-Level Branch | How bands are colored. By First-Level Branch assigns the palette colors (`BAND_PALETTE`) to the first-level branches in turn; deeper branches inherit them. |
| Color | Band color in Single Color mode. Click the swatch to choose it in the standard Color Picker. Grays become K-only colors in CMYK documents. |
| Opacity | Opacity of each band (%). Overlapping bands show through. Defaults to 60%. |

### Buttons

| Item | Description |
| --- | --- |
| Reset | Resets every setting except the flows. Node Height returns to its value when the dialog opened; a font size taken from the labels is kept. |
| Fit to Window | When on, zooms so the previewed diagram fills the given share of the window (65% by default). Cancel restores the original view. |

### Settings Variables

Initial values can be changed under "User settings" at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `DEFAULT_SETTINGS` | — | Initial dialog values (later runs use the last ones) |
| `BAND_PALETTE` | 10 colors | Band colors assigned to the first-level branches in turn (#RRGGBB); converted with the color settings in CMYK documents |
| `LABEL_TINT` | `100` | Color of new labels when there is no template text |
| `OUTLINE_TINT` | `30` | Stroke color of junction shapes and the node box |
| `OUTLINE_STROKE_WIDTH` | `0.75` | Its stroke width (pt) |
| `NODE_HEIGHT_PER_FONT_SIZE` | `16` | Initial height of a text node (multiple of the font size) |
| `NODE_BOX_SIDE_PADDING` | `2` | Side padding of the node box (multiple of the font size) |
| `NODE_BOX_CORNER_RADIUS` | `1.5` | Corner radius of the node box (multiple of the font size) |
| `JUNCTION_PADDING` | `0.6` | Padding between a junction shape and its caption (multiple of the font size) |
| `JUNCTION_POSITION_RATIO` | `0.6` | Where an unplaced junction sits between the start and its children's ends |
| `GROUP_NAME` | `"Sankey Flow"` | Name of the created group |
| `BAND_GROUP_NAME` | `"Bands"` | Name of the band subgroup (the effects are on this group) |
| `JUNCTION_GROUP_NAME` | `"Junctions"` | Name of the junction shape subgroup |
| `LABEL_GROUP_NAME` | `"Labels"` | Name of the subgroup for new labels |

Colors given as K tints become grays of the same darkness in RGB documents.

### Notes

- The key object is detected by actually running the align commands (positions are restored right away).
- Bands leave from the center of the node, so a text node gets a white rounded box behind it to hide them. A shape node with no fill shows the band roots through it.
- The flow field does not accept Tab, so indent with spaces.
- The Outline Stroke effect is applied with the menu command, briefly selecting the band group.
- Settings are stored in `illustrator-scripts/SankeyFlowMaker.json` under `Folder.userData`.

### Update History

- v1.1.2 (2026-10-09) Draw Bands as Strokes is now on by default
- v1.1.1 (2026-10-08) Revised UI wording (Width Scale, junction panel names, etc.); captions enclosed by a circle or rectangle are now center-justified
- v1.1.0 (2026-10-08) Added Draw Bands as Strokes and Fit to Window; Width is now a pair of radio buttons
- v1.0.1 (2026-10-08) Band opacity now defaults to 60%
- v1.0.0 (2026-10-08) Initial version
