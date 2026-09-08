# Draw connectors from the key object to each object

[![Direct](https://img.shields.io/badge/Direct%20Link-AiConnectorBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/AiConnectorBuilder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiConnectorBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Draws a connector from the key object to each of the other selected objects. You choose the route, the stroke, the caps and the arrowheads; the preview on the artboard updates as you change them, and OK keeps exactly what you see. The connectors are collected on a "Connector" layer.

### Main Features

- Five routes: straight, warp, elbow, branch and curve
- Three start points: edge centers, divided and center
- "Share one start point" runs every connector out of one point, picked with a 3x3 widget
- Branch routes are built as one trunk plus branches, so no two branches overlap
- The bend of Elbow and Branch routes is nudged clear of the selected objects
- Stroke width, corner style, dashes (dashed / dotted) and caps (butt / round / projecting)
- Stroke-panel arrowheads (Arrow 8, Arrow 11, dot, circle) picked as icons, with scale, ends and tip placement
- A gap between the end of the connector and the object
- When no key object is set, a dialog picks the start object by its position in a 3x3 grid
- Save and delete the current settings as presets
- Live preview; cancelling discards the connectors it drew
- Remembers the last settings for the session
- Japanese and English UI

### Usage

1. Select two or more objects, then click the one you want as the key object again.
2. Run the script.
3. Choose the shape, stroke and arrowheads, check the preview, and click OK.

If no key object is set, the Start object dialog opens first. Click any cell of the 3x3 grid and the object nearest that position in the selection becomes the start. "Close and pick it manually" closes the dialog so you can set a key object and run the script again. With exactly two objects selected, the left one is used as the start even without a key object.

### Options

**Connector**

| Item | What it does |
| --- | --- |
| Shape | Straight / Warp / Elbow / Branch / Curve |
| Style | Warp style (Bulge / Squeeze) |
| Bend | How much Warp bends and how far Curve bows out (-100 to 100%); a negative value flips the direction |
| Axis | Warp axis. Auto picks one axis for all the connectors together |
| Round corners | Corner radius for Elbow and Branch routes (pt); 0 = square corners |
| Start point | Edge centers / Divided / Center |
| Share one start point | Runs every connector out of the same point on the key object |
| 3x3 widget | Where that shared start point sits; active only while "Share one start point" is on |

**Start point**

| Item | What it does |
| --- | --- |
| Edge centers | Starts at the middle of the key object's edge |
| Divided | Splits the edge into (connectors + 1) parts and spreads the start points along it |
| Center | Aims from the key object's center and starts where that line meets the edge |

With "Share one start point" on, every connector leaves from the same point on the key object, giving a diagram that branches out of a single trunk. The 3x3 widget on the right picks that point.

| Cell | Edge used |
| --- | --- |
| Center | The edge used by the most connectors (auto) |
| Top / Bottom center | The top / bottom edge |
| Left / Right center | The left / right edge |
| Corners | The left or right edge at that height |

**Line**

| Item | What it does |
| --- | --- |
| Stroke width | Line weight (pt) |
| Corner | Miter / Round / Bevel; applies to the bends of Elbow and Branch routes |
| Dashes | None / Dashed / Dotted |
| Divisions | Number of dashes (dashed only); the dash length is solved so both ends finish with a dash |
| Gap | Gap between dashes (dots). For dots it is nudged to the nearest value that lands a dot on both ends |

**Ends & arrowheads**

| Item | What it does |
| --- | --- |
| Shape | Five icons: none / Arrow 8 / Arrow 11 / dot / circle |
| Scale | Arrowhead size in percent, relative to the stroke width |
| Ends | End (the far side from the key object) / Both ends |
| Position | At the end of the path / Beyond the end |
| Gap | Space left between the end of the connector and the object (pt) |
| Cap | Butt / Round / Projecting |

Picking an arrowhead switches the scale, position and cap to that arrowhead's defaults (round cap for the dot and circle, butt for the arrows). Choosing dotted dashes also switches the cap to round, since dots need it to show up.

**Presets**

Save the current settings under a name. Presets are stored in your user settings folder (`AiConnectorBuilder/presets.json`).

### Notes

- At least two objects must be selected.
- The key object is detected by trying the align commands and looking for the object that does not move. The screen is not redrawn during the test, so the shifts are never shown and no undo steps are left behind; the positions are restored.
- Smart Guides are switched off while the script runs and restored on exit.
- Arrowheads cannot be reached from the DOM, so a temporary action is generated and played. The circle is the dot arrowhead with a white circle drawn on top; the white circle follows the scale and is grouped with its line.
- Connectors are created on the "Connector" layer (「コネクター」 in the Japanese UI). An existing layer of that name is unlocked and shown while the script runs, and restored on cancel.
- The live preview is the result. Cancel, ESC or closing the window discards the connectors it drew.
- In Branch routes, the farthest branch on each side of the trunk doubles as the spine, and the nearer branches are plain horizontal (or vertical) lines.
- Setting the warp axis to Auto picks the same axis for every connector, because a per-line axis would mix their appearances.
- The bend of Elbow and Branch routes is moved to the nearest position clear of the selected objects. If that would push it outside the gap, the original position is kept.
- Curve bows the midpoint out at a right angle to the line. The bow is proportional to the line length, so longer connectors curve more; the side it bows to follows the direction of travel, which keeps connectors radiating from the key object turning the same way.

### Original / Acknowledgements

Egor Chistyakov https://x.com/tchegr

### Update History

- v1.0.0 (20260905): Initial release
- v1.0.1 (20260906): Added the Center start point
- v1.0.2 (20260906): Straight is now the default shape; added the start-object dialog for a missing key object, the trunk-plus-branches structure, arrowhead icons, stroke caps, the end gap and "Share one start point"
- v1.0.3 (20260906): Bends are nudged clear of the selected objects; the white circle now follows the arrowhead scale and is grouped with its line; "Share one start point" gained a 3x3 position picker
- v1.0.4 (20260908): Added the Curve shape
- v1.0.5 (20260909): Key object detection no longer adds undo steps
