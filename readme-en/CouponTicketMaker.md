# Ticket and Coupon Shapes

[![Direct](https://img.shields.io/badge/Direct%20Link-CouponTicketMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/CouponTicketMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CouponTicketMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Tear-off lines on tickets and coupons are more tedious to build than they look. You end up spacing perforation dots by hand, adding half-circle notches at both ends of the tear line, rounding the corners — all while keeping the positions and sizes in sync.

This script does the whole set from **a single selected rectangle**. Perforations, zigzag edges, corner treatments and slit / hole cutouts are combined, and you tune the numbers against a live preview.

## How to use

1. Select **exactly one** rectangular path.
2. Run the script.
3. Adjust the settings in the dialog (the preview updates as you type).
4. Click **OK** to apply the result to the object you selected.

The script stops in these cases:

| State | Message |
| --- | --- |
| No document open | Please open a document. |
| Nothing selected | Please select a rectangle. |
| More than one object selected | This script cannot run with multiple selections. Please select only one object. |
| A group is selected | This script cannot run when a group is selected. Please select a single object. |
| Not a rectangle | Please select exactly one rectangular path. |

A path counts as a rectangle when it has four anchor points and all four sit on the corners of its bounding box (0.01 pt tolerance). Rotated rectangles and rounded rectangles are not accepted.

## Presets

The dropdown at the top of the dialog holds ten presets.

| Preset | Contents |
| --- | --- |
| 0: Clear | Everything off — a clean starting point |
| 1: Double Rounded + Divider | Divider line plus double rounding |
| 2: Two Side Holes | Circular holes on the left and right edges |
| 3: Perforation | Dotted divider line plus rounded corners |
| 4: Inverse Round Corners | Inverse arcs carved out of the four corners |
| 5: Triangular Side Slits | Triangular slits on the sides plus rounded corners |
| 6: Perforation + T/B Zigzag | Side perforations plus zigzag on the top and bottom |
| 7: Chamfer + Dashed Divider | Chamfered corners plus a dashed divider line |
| 8: Divider + Circle Edges + Zigzag | Divider line, circular edges and L/R zigzag |
| 9: Dashed Divider + Triangle Edges + Left Hole | Dashed divider, triangular edges and a hole on the left only |

**Save** names the current settings and adds them to the dropdown. That copy lives **only for the current session** and disappears when the script closes.

The save dialog also prints a ready-to-paste code snippet. Append it to the `PRESETS` array inside the script and the preset becomes permanent.

Preset names are localizable. Rewrite `name: "..."` in the snippet as `name: { ja: "...", en: "..." }` to get a name per locale (a plain string keeps working).

## Corner

How the four corners are treated.

| Option | Behaviour |
| --- | --- |
| None | Leave the corners alone |
| Rounded | Apply the *Round Corners* live effect |
| Inverse Round | Carve the corners out with reversed arcs |
| Chamfer | Cut the corners off at 45° |

*Inverse Round* works by stroking the top and bottom edges with a dash whose weight equals the size, outlining that stroke so only squares at the corner positions remain, and subtracting the result. The dash gap is deliberately large (1000 pt) so nothing appears between the corners.

## Zigzag

Adds a saw-tooth cutout along the edges — *L/R* on the left and right, *T/B* on the top and bottom.

| Field | Behaviour |
| --- | --- |
| Size | The **diagonal** of a single tooth (each one is a square rotated 45°) |
| Repeat | Number of teeth per edge |
| Gap | Distance between teeth. Enabled **only when Repeat is 2 or more** |

The run is centred on each edge. Nothing clamps the total length, so you can deliberately let the teeth overshoot the edge.

*L/R* zigzag and the side perforation below use the same edges and cannot be combined. Turning the perforation on resets the zigzag to *None*.

## L/R

### Perforation

Runs a dotted perforation along the left and right edges.

| Field | Behaviour |
| --- | --- |
| Enable | Build the side perforations |
| Link to Split Line | Copy weight, gap and inset length from the divider line |
| Weight | Diameter of each dot |
| Gap | Distance between dots |
| Inset Length | How far the line pulls back from each end (**0 or less only**) |

*Link to Split Line* only takes effect while the divider line is set to **Dot**. While linked, the weight, gap and inset fields are dimmed and the divider values are used instead.

*Inset Length* is how far both ends are pulled inwards. Positive values are rejected and snap back to 0.

## Slit / Hole

Cuts a notch into the middle of the left and right edges.

| Field | Behaviour |
| --- | --- |
| None / Circle / Triangle | Shape of the notch |
| Left / Right | Which edge to cut (nothing happens if both are off) |
| Size | Diameter for a circle, diagonal for a triangle |

The shape is centred on the edge, so only half of it (a semicircle or a triangle) is visible in the result.

## Center Divider

Everything around the tear line that splits the ticket in two. Turning *Enable* off dims the divider line and edge settings as well.

The number field and slider next to *Enable* set the position of the divider line as an offset from the centre. The slider is limited to **half the width** of the selected rectangle, so the line can never leave the shape.

### Divider Line

| Field | Behaviour |
| --- | --- |
| Dot | Round perforation |
| Dash | Straight dashed line |
| Weight | Dot diameter / dash thickness |
| Gap | Distance between dots or dashes |
| Inset Length | How far the line pulls back from each end (**0 or less only**) |

With **Dash** selected and a non-zero inset length, the script **picks the dash count so that the dash length lands as close to the gap as possible**. Neither end is cut off mid-dash, which is what makes it read as a tear line. When the inset length is 0, dashes are simply repeated at the same length as the gap.

### Edge

Decoration placed where the divider line meets the top and bottom edges — typically a half-circle notch at each end of the tear line.

| Field | Behaviour |
| --- | --- |
| None / Circle / Triangle | Shape of the edge |
| Double Rounded | Split the rectangle at the divider line and round each half |
| Size | Diameter for a circle, diagonal for a triangle |
| Edges Only | Place the edge shapes without drawing the divider line |

**Choosing an edge shape recalculates the divider line's inset length automatically**, pulling the line back by roughly half the edge so that the line and the notch join cleanly. The multiplier depends on the shape.

| Edge | Value written into Inset Length |
| --- | --- |
| Circle | Size × 1.0, negated |
| Triangle | Size × 1.2, negated |
| Double Rounded | Size × 2.0, negated |

*Double Rounded* is the odd one out: instead of adding a shape, it **rebuilds the rectangle itself**. The rectangle is split at the divider line, each half is rounded, and the two halves are merged again — so the ticket pinches inwards on both sides of the tear line. While it is on, the Corner panel is dimmed and the edge shape is forced to *None* (switching it back off restores the corner mode to *Rounded*).

*Edges Only* dims the divider line panel and places just the edge shapes, with no line between them.

## Preview and applying

While *Preview* is on, the result is built on a dedicated layer named **プレビュー / Preview** and the **original object is hidden temporarily**. The original is never touched, so you can experiment freely.

**OK** clears the preview layer and re-runs the same operations on the original object. **Cancel** leaves nothing behind.

*Expand Appearance* runs Object → Expand Appearance after applying, converting live effects such as rounded corners into real paths.

*Outline View* at the bottom left toggles Illustrator's view mode — useful for checking where the perforations landed.

## Units

Every value uses the **ruler unit** (Preferences → Units → General). The one exception is the divider line's *Weight*, which follows the **stroke** unit. The current unit is shown to the right of each field.

Number fields respond to the arrow keys.

| Key | Step |
| --- | --- |
| Up / Down | 1 |
| Shift + Up / Down | 10 (snaps to multiples of 10) |
| Option + Up / Down | 0.1 |

## Notes

Every generated shape is **subtracted from the original rectangle** (the *Pathfinder → Minus Front* live effect). Perforations, zigzag teeth and holes are all built as cutters first and removed from the rectangle in one pass at the end.

Because of that, the original object's appearance is **normalised down to a fill**.

| Original | After normalising |
| --- | --- |
| Has a fill | The stroke is dropped, the fill is kept |
| Stroke only | The stroke colour is moved to the fill |
| Neither | Filled with 60% black (CMYK) or 60% gray |

Dotted perforations only come out right when the stroke alignment is centred, and Illustrator's DOM cannot set stroke alignment. The script therefore loads a temporary action (`StrokeDot`) at run time and applies it. The action is written to `~/StrokeDot.aia`, deleted immediately, and unloaded when the script finishes.

While the dialog is open, anchor-point display and the bounding box are toggled temporarily and restored on exit.

### note

- [Article (Japanese) | DTP Transit 別館](https://note.com/dtp_tranist/n/n2e949946228a)
