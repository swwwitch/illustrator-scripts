# Path Optimization

[![Direct Link](https://img.shields.io/badge/Direct%20Link-PathCleanupTool.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/PathCleanupTool.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathCleanupTool.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Traced paths, paths brought in from another application, paths edited over and over. They may look fine, but turn on the anchor points and you often find a pile of "why is there a point here?" — anchors spaced evenly along a straight run, points stacked at exactly the same coordinates, handles still pulled out on a segment that is perfectly straight.

Deleting them changes nothing visually. That is precisely what makes them awkward: **you cannot see whether they are safe to delete**. And selecting them one by one by hand is not realistic either.

So this script shows you **what and how many will go** before it runs. Every time you toggle a checkbox or move a tolerance, the anchor and handle counts update as "before → after".

<img alt="" src="" width="50%" />

## Usage

1. Select the paths (groups and compound paths are fine as they are)
2. Run the script
3. Adjust the checkboxes and tolerances while watching the estimates in the info panel
4. [OK]

The dialog has two tabs, Removal Targets and Transform. **Only the processing of the tab that was active when you pressed OK is run.** The settings of both tabs are never applied together.

## Info panel

Always visible at the top, it shows the anchor point and handle counts as "current → after". The path count gets the arrow only on the Transform tab, where it can change (nothing on the Removal Targets tab changes how many paths there are).

The estimate is calculated without touching the DOM. The target paths are copied into a lightweight model (arrays of coordinates) and counted there. So you get a preview before running, and the document is not modified at all.

**The preview and the actual processing call the same functions in the same order.** The real run also computes on that same model first, then writes the result back to the path. There is no separate algorithm for previewing and for running, so the numbers on screen cannot drift from what you get.

The one exception is Fill holes: the result of the Pathfinder cannot be calculated accurately in advance, so its estimate shows `-`.

## Removal Targets tab

| Item | What is removed |
| --- | --- |
| Duplicate anchor points | Merges **consecutive** points at the same coordinates into one (identical coordinates further apart are not affected) |
| Collinear anchor points | Points that are collinear with their neighbours and have no handles pulled out |
| Handles on straight segments | Returns the handles of a segment that can be treated as straight to the anchor position (no visible change) |

All three are on by default.

### The two tolerances

"Collinear anchor points" and "Handles on straight segments" each have their own tolerance (0.01–3.00). The number field and the slider are linked, and moving either updates the estimates in the info panel.

The larger the value, the looser the "treat as straight" test, and the more is removed. They share the label "Tolerance", but they measure different things.

- **For anchors**: the cross product of three points decides whether they are collinear
- **For handles**: the perpendicular distance (pt) from the line joining the two end anchors to the tip of the handle

Clearing a checkbox dims the corresponding tolerance row.

### Why the execution order goes back and forth

The three operations do not simply run once each in the order listed. The actual order is:

```
duplicate anchors → redundant anchors → handles → redundant anchors → duplicate anchors
```

Removing the handles from a straight segment turns anchors that were previously judged "cannot be removed, it has handles" into new candidates. Hence the turn-around that puts anchor removal through a second time.

Only one place in the script holds this order, and the info panel's estimate calls into it.

### Endpoints are never removed

The endpoints of an open path are excluded from both anchor removal and duplicate removal. Handle cleanup likewise leaves alone the start-side handle of the first segment and the end-side handle of the last segment of an open path. This is to keep the shape of the path from breaking.

## Transform tab

It is divided into three groups, but **all six radio buttons form a single exclusive group** (only one can be selected).

### Convert anchor points

| Mode | What it does |
| --- | --- |
| To smooth points | Makes every anchor a smooth point and gives it handles |
| To corner points | Makes every anchor a corner point and removes its handles |

"To smooth points" does not simply attach handles of a uniform length. One third of the distance to the neighbouring anchors is used as the base length, and two corrections are applied on top of it.

- **Angle correction**: the sharper the turn, the shorter the handle (up to 35% shorter)
- **Balance correction**: the more lopsided the distances to the neighbouring anchors, the shorter the handle (up to 25% shorter)

The direction is the tangent derived from the neighbouring anchors, used as is — it is never rounded to 8 directions or the like.

### Add anchor points

| Mode | What it does |
| --- | --- |
| At midpoints | Adds one point at the middle of each segment |
| Add Extreme Points | Adds points where the curve's tangent is horizontal or vertical (the extrema) |

"At midpoints" calls Illustrator's own Add Anchor Points menu command.

"Add Extreme Points" is a custom implementation. Each segment is treated as a cubic bezier, the parameter t where the derivative becomes 0 is solved for, and the curve is split with de Casteljau's algorithm. **The shape of the original curve does not change** (splitting leaves the bezier's trajectory identical). Straight segments are not affected.

Adding extrema does not rebuild the path: the point count of the same PathItem is increased and all points are overwritten, so the appearance is preserved.

### Other

| Mode | What it does |
| --- | --- |
| Split at anchor points | Splits each segment into an independent open path (the original path is deleted) |
| Fill holes | Releases the compound path, unites it, and fills the holes |

"Split at anchor points" carries the original path's fill and stroke (colour and weight) over to each segment. The resulting paths are created inside the same group as the original (the children of a compound path are the exception: an open path cannot go back into a compound path, so they move up one container).

"Fill holes" is a run of consecutive menu commands: group → release compound path → live Pathfinder (unite) → expand appearance → ungroup. **If the selection contains no hole (a compound path with two or more subpaths), the radio button itself is dimmed and cannot be selected.**

## Targets

PathItem, CompoundPathItem, GroupItem (PathItems are collected recursively from all of them)

Locked and hidden objects are skipped automatically. The test walks up the parents, so **even if the object itself is not locked, it is excluded when the group or layer it belongs to is locked or hidden**.

## Notes

This script fixes the selection as an array at the moment the dialog opens, and reuses it from then on. Both the estimate in the info panel and the actual processing after OK look at the same target list, so the display and the execution never drift apart even if the selection changes while the dialog is open.

Just before processing, the selection is restored from that snapshot. "At midpoints" and "Fill holes" use menu commands and therefore depend on the selection state; for those, only the PathItems left after skipping are reselected. If not a single target can be restored, the script says so and stops.

The Removal Targets tab does not rebuild the paths: it keeps the same PathItem, drops the surplus points and overwrites the rest (the same approach as "Add Extreme Points"), so the appearance survives. Only paths that actually changed are written back — a path with nothing to remove is left untouched.

Using an index (`selection === 0`) to detect the tab can misfire across ScriptUI environments, so the branch compares against the tab object itself.

The dialog position is remembered only while Illustrator is running (position only — restoring the size as well can leave the dialog too small to read).

## Change log

- v1.6.0 (2026-07-31) Unified the preview and the actual processing onto a single algorithm (current version). "Split at anchor points" now places the results inside the original group
- v1.5.2 (2026-07-14) Added the "Add Extreme Points" mode
- v1.5.1 (2026-03-20)
- v1.0 (2026-03-01) Initial release

### note

- [Optimizing paths in Illustrator (Japanese)](https://note.com/dtp_tranist/n/nd82f59bf63a8)
