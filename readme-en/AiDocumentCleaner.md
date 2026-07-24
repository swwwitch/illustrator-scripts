# AiDocumentCleaner.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-AiDocumentCleaner.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiDocumentCleaner.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiDocumentCleaner.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

A cleanup tool that removes unneeded elements from a document in one pass. Choose the targets in the dialog and tidy the whole document at once.

Swatches, symbols, brushes, and graphic styles are pruned by building and playing each panel's "Select All Unused -> Delete" action, so only items Illustrator itself judges unused are removed. Paths/objects, groups/layers, guides, and artboards can be cleaned in the same run.

## Main features

- **Panel items**: swatches, symbols, brushes, graphic styles, paragraph styles, character styles
  - Swatches / symbols / brushes / graphic styles are pruned via an action (unused only)
  - Paragraph / character styles have no usage info, so all but the default (first) are removed only in force mode
- **Paths / Objects**: stray points, empty text frames, unpainted (invisible) paths, 0% opacity, hidden objects, broken-link placed images
- **Groups / Layers**: recursively removes empty groups, empty layers, and sublayers
- **Guides** (pick one): Guides on unlocked layers / All guides (force) / Outside the active artboard
- **Artboards**: objects outside all / the active artboard, and empty artboards
- "Delete items even if in use" (force) removes everything but protected/default items
- Reports the deleted count per type (zero-count types are omitted)
- Japanese and English UI

## Usage

1. Open the document you want to clean up.
2. Run `AiDocumentCleaner.jsx`.
3. Check the targets you want to delete in the dialog.
4. To go beyond unused items and remove everything but protected/default ones, enable "Delete items even if in use" (force).
5. Pick one guide option with the radio buttons (default is "Don't delete").
6. Click Delete to run; the per-type deleted counts are shown.

## Targets

| Category | Targets |
| --- | --- |
| Panel items | Swatches / symbols / brushes / graphic styles / paragraph styles / character styles |
| Paths / Objects | Stray points / empty text frames / unpainted paths / 0% opacity / hidden objects / broken-link placed images |
| Groups / Layers | Empty groups / empty layers and sublayers |
| Guides | Guides on unlocked layers / All guides (force) / Outside the active artboard |
| Artboards | Objects outside all / the active artboard / empty artboards |

## Notes

- Deletion-prone options (hidden objects, broken-link placed images, objects outside all / the active artboard) start unchecked.
- Guides default to "Don't delete".
- Paragraph and character styles have no usage info, so only force mode removes all but the default.
- Group / layer cleanup runs after the other deletions, so a parent emptied by path/object removal is cleaned in the same pass.
- The temporary action is always unloaded and deleted after the run.
- Clipping paths, compound-path members, and guides are excluded from the "unpainted paths" option.
- Guides are only removed by the Guides section (the "unpainted paths" and outside-artboard options never touch them).
- Back up your file before running.

## Article

[Remove unneeded document elements in one pass with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n0d70178f0f65)

## Changelog

- v1.0.3 (2026-07-24): Run empty group/layer cleanup after the other deletions, fix missed items in hidden-object removal, exclude guides from the artboard emptiness check, and move the temp file to the OS temp folder. Also exclude guides from outside-artboard deletion and recurse empty-group cleanup into sublayers. Clarified labels ("Guides on unlocked layers", "Outside the active artboard", "Paths with no fill or stroke") and added a divider above the force option
- v1.0.2 (2026-06-27): Initial release
