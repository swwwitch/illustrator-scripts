# Dash Calculator (DashGapCalculator)

[![Direct](https://img.shields.io/badge/Direct%20Link-DashGapCalculator.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/DashGapCalculator.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashGapCalculator.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Calculates the dash length from the length of the selected path, the number of segments, and the gap, then applies it as a dashed stroke.
- You can also solve the gap from a given dash length, or fill the path with a random dash pattern.
- The dash offset (phase), stroke cap and path direction can be set from the same dialog.
- Everything is previewed live, and Cancel restores the original state.

This is for the moments when you want a dotted line to close perfectly in five steps, or a dash break to land exactly on a corner, without doing the arithmetic yourself.

### How to use

1. Select the path (open or closed) you want to dash. Multiple selections are supported.
2. Run the script.
3. Adjust the settings in the dialog (the preview updates as you change values).
4. Click OK to confirm.

The script stops in these cases:

| State | Message |
| --- | --- |
| No document open | No document is open. |
| Nothing selected | Select one path (open or closed). |
| The selection contains no path | Select exactly one path (open or closed). |

Only **paths** are processed. Text frames, symbols and the like are simply ignored when they are part of the selection.

### Dialog settings

| Item | Description |
| --- | --- |
| Selected Path Info | Length of the first path. The count in parentheses appears when several paths are selected |
| Segments | How many parts the path is divided into (integer, 1 or greater) |
| Gap | Length of the empty space |
| Dash | Length of the dash |
| Calculation | Gap→Dash / Dash→Gap / Random |
| Offset | Dash offset (phase). Enabled by the checkbox |
| Partial Display | Shows only one of the divided dashes |
| Cap | Butt / Round / Projecting |
| Adjust ends | On an open path, distributes the dashes so both ends finish with a dash |
| Reverse Path Direction | Reverses the path direction (which moves where the dashes start) |

The input field and the result label share the same spot and swap depending on the calculation mode: in Gap→Dash the gap is editable and the dash is the result, in Dash→Gap it is the other way round.

### Calculation

#### Gap→Dash

Enter the gap and the dash length is calculated.

#### Dash→Gap

Enter the dash length and the gap is calculated.

In both modes, values that would make the result zero or negative cannot be confirmed; an alert reports the maximum value you can use.

#### Random

Builds a random pattern of six entries — three dashes and three gaps.

- The dashes are random; the gaps all use the value you entered.
- **Each path gets its own random draw**, so selecting several paths gives each of them a different dash pattern.
- Clicking the radio button generates a new pattern, so click again if you do not like the result.
- Segments, Dash, Partial Display and Adjust ends are dimmed (they are not used in this mode).
- With the Butt cap, a per-unit minimum (mm = 1 / pt = 2 / Q, H = 4 / others = 2) keeps the dashes from disappearing. Round and Projecting caps have no minimum.

### Adjust ends

An option for open paths; it is dimmed when a closed path is selected.

| State | How one cycle is measured |
| --- | --- |
| On (open path) | Dashes are distributed so both ends finish with a dash: the dash count equals Segments, and the gap appears (Segments − 1) times |
| Off, or closed path | One cycle (dash + gap) = path length ÷ Segments. The tail may end mid-cycle |

With Adjust ends on and Segments set to 1, the dash covers the whole path.

### Offset

Turn on the checkbox to set the dash offset (phase). Besides typing a value, you can pick a preset expressed as a fraction of one cycle (dash + gap).

| Preset | Offset |
| --- | --- |
| 1/4 | one cycle × 0.25 |
| 1/2 | one cycle × 0.5 |
| 3/4 | one cycle × 0.75 |

When a typed value matches one of them, the matching preset is selected automatically.

### Partial Display

Shows only one of the divided dashes and hides the rest — handy when you want to reveal just a portion of a path.

Turning it on sets the gap to 0, and the dash length becomes "path length ÷ Segments" (or the distributed length when Adjust ends is on for an open path). It cannot be combined with Random.

### Clear Dashes

**Clear Dashes** previews the path with its dashes removed. Click OK to confirm, or change any value to return to the normal preview.

### Units

All values use the unit set in Preferences → Units → Stroke. The Q / H label follows the East Asian options in the preferences.

Numeric fields accept the arrow keys.

| Key | Step |
| --- | --- |
| Up / Down | 1 |
| Shift + Up / Down | 10 (snaps to multiples of 10) |
| Option + Up / Down | 0.1 |

### Multiple selections

The length and the results shown in the dialog belong to the **first** path, but each path is calculated from its own length when the dashes are applied. Paths of different lengths are therefore each divided into the number of segments you asked for.

### Notes

- Cancel (and the close button / Esc) restores the dashes, cap and path direction to the state they had before the dialog opened.
- Reverse Path Direction runs the same command as Object → Path → Reverse Path Direction. If you toggled it in the dialog, Cancel reverts it.
- Dashes are measured from the start point of the path. If they do not begin where you expect, adjust the offset or reverse the direction.
- Segments, gap, dash, offset, cap, calculation mode, Adjust ends, Reverse Path Direction and the random pattern are remembered and restored the next time you run the script.

### Changelog

- v1.4 (2026-02-25) : Initial version
- v1.5 (2026-02-25) : Support for multiple paths, partial display and offset presets
- v2.0 (2026-02-28) : Added the random mode
- v2.0.1 (2026-08-13) : Reorganized the internal structure; fixed the random mode being blocked by the unused segment count; added tooltips to the dialog

### Script info

- Version: v2.0.1
- First release: 2026-02-25
- Last updated: 2026-08-13
- Article: https://note.com/dtp_tranist/n/n868bedb96542
