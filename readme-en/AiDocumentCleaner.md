# AiDocumentCleaner.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-AiDocumentCleaner.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiDocumentCleaner.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiDocumentCleaner.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

A cleanup tool that removes unneeded elements from a document in one pass. Choose the targets in the dialog and tidy the whole document at once.

Swatches, symbols, brushes, and graphic styles are pruned by building and playing each panel's "Select All Unused -> Delete" action, so only items Illustrator itself judges unused are removed. Paths / objects, groups / layers, guides, and artboards can be cleaned in the same run.

The run can target the frontmost document, every open document, or all .ai files in a chosen folder.

## Main features

- **Scope** (pick one)
  - Frontmost document (default)
  - All open documents (the label shows the current count; nothing is saved)
  - Folder: opens each .ai file directly inside the folder, cleans it, **saves over the original**, and closes it
- **Panel items**: swatches, symbols, brushes, graphic styles, paragraph styles, character styles
  - Swatches / symbols / brushes / graphic styles are pruned via an action (unused only)
  - Paragraph / character styles have no usage info, so all but the default (first) are removed only when "Delete panel items even if in use" is on
- **Paths / Objects**: stray points, empty text frames, unpainted (invisible) paths, 0% opacity, hidden objects, broken-link placed images
- **Groups / Layers**: recursively removes empty groups, empty layers, and sublayers
- **Guides** (pick one): Clear Guides (locked ones remain) / All guides (unlocks everything) / Outside the active artboard
- **Artboards**: objects outside all / the active artboard, and empty artboards
- "Delete panel items even if in use" removes panel items even when the document uses them (protected and default items remain; artboards are unaffected)
- Reports the deleted count per type (zero-count types are omitted; a batch run reports the totals across all files)
- A Preset popup (Default / All off / All on / Panel items only / Custom), plus option/alt-click to toggle a whole panel at once
- Remembers the previous settings and the dialog position for the session only (discarded when Illustrator quits; nothing is written to preferences)
- Japanese and English UI

## Usage

1. Open the document you want to clean up (not needed for the folder target).
2. Run `AiDocumentCleaner.jsx`.
3. Pick the scope in the Scope panel at the top. For the folder target, choose a folder with the Choose... button.
4. Check the targets you want to delete.
5. To remove panel items the document still uses — not just the unused ones — enable "Delete panel items even if in use".
6. Pick one guide option with the radio buttons (default is "Don't delete").
7. Click Run to start; the per-type deleted counts are shown.

## Choosing the scope

| Target | Behavior | Saved |
| --- | --- | --- |
| Frontmost document | Cleans the document you are working on | No |
| All open documents | Cleans every open document in turn | No |
| Folder | Opens and cleans each .ai file directly inside the folder | **Saved over the original** |

Once a folder is chosen, the radio label shows how many .ai files the run will cover, e.g. "Folder (13)", so a zero is obvious before you run.

The chosen folder's path is shown on the line below the Choose... button. A folder under Dropbox drops the mount-path prefix; anything else is shortened with `~` for the home folder (hover to see the absolute path).

## Targets

| Category | Targets |
| --- | --- |
| Panel items | Swatches / symbols / brushes / graphic styles / paragraph styles / character styles |
| Paths / Objects | Stray points / empty text frames / unpainted paths / 0% opacity / hidden objects / broken-link placed images |
| Groups / Layers | Empty groups / empty layers and sublayers |
| Guides | Clear Guides (locked ones remain) / All guides (unlocks everything) / Outside the active artboard |
| Artboards | Objects outside all / the active artboard / empty artboards |

## Presets and bulk toggling

The Preset popup at the left of the button area switches every deletion target at once.

| Item | Behavior |
| --- | --- |
| Default | Restores the initial state (the four deletion-prone options off, guides on "Don't delete") |
| All off | Clears every checkbox and sets guides to "Don't delete" |
| All on | Checks everything and sets guides to "All guides (unlocks everything)" |
| Panel items only | Checks the panel items and nothing else (guides go back to "Don't delete") |
| Custom | Shown automatically once you change something by hand; picking it does nothing |

The force option is never touched by a preset — being destructive, it stays manual.

Option/alt-clicking a checkbox sets every checkbox in that panel to the clicked one's state.

## Notes

- **The folder target saves over your original .ai files and cannot be undone.** A confirmation dialog showing the file count and folder path appears before the run. Back up your files first.
- The folder target only covers files directly inside the folder; subfolders are not included.
- During a folder run, alerts (missing fonts and the like) are suppressed so the batch doesn't stall, and the setting is restored afterwards. A file that fails is closed without saving and its name is listed in the summary.
- "All open documents" never saves. Review the results and save yourself.
- Deletion-prone options (hidden objects, broken-link placed images, objects outside all / the active artboard) start unchecked.
- Guides default to "Don't delete".
- "All guides (unlocks everything)" temporarily unlocks locked layers, sublayers, and the guides themselves, then restores every lock afterwards.
- Paragraph and character styles have no usage info, so they are removed only when "Delete panel items even if in use" is on.
- Group / layer cleanup runs after the other deletions, so a parent emptied by path/object removal is cleaned in the same pass.
- The temporary action is always unloaded and deleted after the run.
- Clipping paths, compound-path members, and guides are excluded from both "unpainted paths" and "0% opacity" — removing one on its own would break the structure it belongs to.
- Pruning unused panel items plays an action recorded on a Japanese UI. Other UI languages are unverified; if the action cannot be played, the summary says so.
- Swatch groups left empty after swatches are removed are cleared out too (they are not counted).
- Guides are only removed by the Guides section (the "unpainted paths" and outside-artboard options never touch them).
- The previous settings, target folder, and dialog position are kept until Illustrator quits, then reset (via a persistent engine declared with `#targetengine`). A remembered folder that no longer exists is not restored.
- Back up your file before running.

## Article

[Remove unneeded document elements in one pass with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n0d70178f0f65)

## Changelog

- v1.1.0 (2026-08-02): Added a Scope panel for choosing what to process (frontmost document / all open documents / a folder). The folder target cleans every .ai file directly inside the folder and saves over the originals, with a confirmation dialog before the run. The chosen folder's path is displayed (Dropbox mount prefix dropped, otherwise shortened with `~`). "All guides (unlocks everything)" now also reaches guides on locked sublayers and guides that are locked themselves. Items whose action could not be played are now reported as a failure instead of a zero count. Added a Preset popup (Default / All off / All on / Panel items only / Custom) and option/alt-click to toggle a whole panel, plus session-only memory of the previous settings, target folder, and dialog position. Fixed the force option removing non-empty artboards: it now applies to panel items only, and is relabelled "Delete panel items even if in use". Fixed empty-layer removal leaving a layer unlocked when the removal failed. Reworked the UI wording (title "Clean Up Documents", "Delete items even if in use" to "Delete panel items even if in use", the guide options, the Choose... and Run buttons) and added tooltips to the Choose... and Run buttons, the "Don't delete" guide option, and the Scope panel
- v1.0.3 (2026-07-24): Run empty group/layer cleanup after the other deletions, fix missed items in hidden-object removal, exclude guides from the artboard emptiness check, and move the temp file to the OS temp folder. Also exclude guides from outside-artboard deletion and recurse empty-group cleanup into sublayers. Clarified labels ("Guides on unlocked layers", "Outside the active artboard", "Paths with no fill or stroke") and added a divider above the force option
- v1.0.2 (2026-06-27): Initial release
