# SelectionInspector.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-SelectionInspector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SelectionInspector.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionInspector.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

"How many text frames are in this document?" "How many anchor points?" "Are there any broken links left?" — the numbers you want before handing off a job, or when sizing up how much work is left, aren't all in Illustrator's panels. The Document Info panel has some of them, but never side by side with the current selection.

This script is a **persistent palette that shows the selection and the whole document next to each other**. Keep it open, change the selection, press Refresh, and the numbers update. Object notes (the Attributes panel note field) can be read and edited from the same palette.

## Usage

1. Run the script (the palette appears)
2. Select the objects you want to count
3. Press Refresh (Cmd+R)

With nothing selected, only the document totals are counted. The palette stays open, so changing the selection and pressing Refresh is all it takes to get new numbers.

Running the script again while the palette is open closes the existing one first, so you never end up with duplicates.

## What is counted

Every value is formatted as **"Selection / All"** (Artboards shows the document total only).

The palette has two columns.

### Left column

| Panel | Rows |
| --- | --- |
| Basics | Artboards (total only), Objects |
| Text Frames | Text frames, point type, area type, path text |
| Characters & Paragraphs | Characters, paragraphs, line breaks |
| Images | Linked, embedded, broken links |
| Notes | The selected object's note (read-only) |

### Right column

| Panel | Rows |
| --- | --- |
| Groups | Groups, clipping groups |
| Transparency | Opacity < 100, blend mode != Normal |
| Paths | Paths, open paths, closed paths, anchor points, handles, compound paths, compound shapes |
| Guides | Ruler guides, artboard guides, other guides |

A status line at the bottom reports the result ("Selected 3 object(s)") or any error.

## Counting rules

A few rules are worth knowing.

- **Object count (All)** includes the contents of groups and compound paths.
- **Text and path stats** are counted recursively inside groups, so selecting a single group still counts the text and paths within it.
- **Guides** (paths with `guides=true`) are excluded from the path stats; they are counted separately in the Guides panel.
- **Line breaks** is the number of newline characters in the text contents.
- **Handles** counts direction lines that sit away from their anchor point (up to two per anchor).
- **Broken links** are linked images whose file is missing. They are also included in the linked count.
- **Compound shapes** are detected as `PluginItem`s whose name contains "Compound Shape".

### How guides are classified

Guides are split into three kinds by their shape and the artboard sizes. Only horizontal or vertical two-point paths can be classified.

| Kind | Test |
| --- | --- |
| Ruler guide | Longer than the largest artboard edge in the document |
| Artboard guide | Matches the width or height of some artboard (0.5pt tolerance) |
| Other guide | Neither of the above |

Guides made from shapes, and anything else that isn't a two-point path, fall into "Other guides".

## Editing notes

The Notes tab edits the note (Attributes panel note field) of each selected object.

- One field per selected object.
- Fields are ordered by position: top to bottom, then left to right at the same height.
- Apply writes the note to that object and recounts automatically.

The Notes panel on the Info tab shows the note itself when there is exactly one, and "Multiple notes found" when there are several.

If the selection changed after the palette was refreshed, Apply reports "Selection changed. Please refresh." and writes nothing.

## Export

Export saves the stats as a plain text file on the desktop.

- File name: `count-<document>-YYYYMMDD.txt`
- Contents: the same rows as the palette, grouped into sections, plus any non-empty notes.

## Keyboard shortcuts

| Key | Action |
| --- | --- |
| Cmd+R | Refresh (recount) |
| Opt+I | Info tab |
| Opt+M | Notes tab |
| Esc | Close |

Enter is deliberately unassigned, because it conflicts with editing notes.

## Notes

- Because the palette runs in its own engine (`#targetengine`), it cannot touch the document directly. Counting and note writing are delegated to the main engine via BridgeTalk.
- Selection changes are not detected automatically — press Refresh (Cmd+R).
- With no document open, the status line says so and every value shows "-".
- If Illustrator doesn't answer within 10 seconds, the status line reports "No response from Illustrator".
- The palette position is kept for the current Illustrator session only and resets on restart.

## Article

[Illustrator: a script for inspecting the current selection (Japanese)](https://note.com/dtp_tranist/n/nefcb1ce828ce)

## Changelog

- v1.7.0 (2026-07-02): Palette conversion (`#targetengine` + BridgeTalk delegation), refresh button and status line, localization reorganized into categories with `L()`
- v1.6.1 (2026-03-16): Localization cleanup (export strings and section headings moved into `LABELS`)
- v1.6 (2026-03-16): Added the Notes panel below Images
- v1.5.3 (2026-03-14): Remember and restore the dialog position during the session
- v1.5.2 (2026-03-12): Recursively count text stats inside groups
- v1.5.1 (2026-03-12): Recursively count path stats inside groups
- v1.5 (2026-03-02): Added the guide breakdown, excluded guides from path stats, updated the layout and export
- v1.4.1 (2026-03-02): UI tweaks (OK to Close, removed Cancel, moved Export to the left)
- v1.4 (2026-03-01): Added the forced line break (soft return) count
- v1.3 (2026-03-01): Added the handle count for paths
- v1.2 (2025-08-07): UI adjustments
- v1.1 (2025-08-06): Added the export feature
- v1.0 (2025-08-06): Initial version
