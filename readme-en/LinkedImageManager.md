# Linked image management

[![Direct Link](https://img.shields.io/badge/Direct%20Link-LinkedImageManager.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/LinkedImageManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LinkedImageManager.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Checking linked images before handing off artwork means answering the same questions every time. Are any links broken? Is the resolution high enough? Did an RGB image slip in? How many places use the same file?

Illustrator's Links panel can answer each of those, but only one image at a time — there is **no way to compare them side by side**. Checking PPI means selecting an image, expanding the panel, selecting the next image, expanding again. Sorting by "lowest PPI first" is not possible at all.

This script analyzes every placed image at once and shows them **as a sortable, filterable table** in a persistent palette. Relinking, renaming, and deleting can all be done from the same list. **Embedded images appear in the same table**, and switching them back to links (unembedding) — or the other way around (embedding) — happens in the same place.

## Usage

Running the script opens a palette listing the placed images (PlacedItem) and the embedded images (RasterItem) in the document.

If a placed image was selected before running, the matching row is selected from the start. The canvas view does not move in that case.

The palette also opens with no document open. Open a document afterwards and click Refresh to load it. When there are no placed images, or no document is open, the palette collapses to a simplified view with the options and list hidden.

The palette can also be closed with the `Esc` key.

## Reading the list

The leading icon column and the file name column are always visible. Every other column is toggled with the checkboxes in the Columns panel.

| Column | Contents |
| --- | --- |
| Icon | ✓ Linked / ⚠ Missing / ⟳ Needs update / ▣ Embedded |
| File name | Linked file name |
| Size | File size |
| Uses | How many places the same file is placed in |
| Width (mm) / Height (mm) | Actual placed dimensions |
| Scale | Scaling applied when placed |
| PPI | Effective resolution |
| Color space | Color mode (plus ICC profile name) |
| Artboard | Which artboard the image sits on |

### Rows for embedded images

Embedded images are listed as "▣ Embedded". They have no source file on disk, so **the file path, size, and PPI show `---`**. The file name comes from the item name (which usually keeps the original file name after embedding), and the color space comes from `imageColorSpace`.

Rasterized artwork also appears as an embedded image. An embedded PSD that has been ungrouped is listed once per layer.

### How status is determined

"⟳ Needs update" compares the link timestamp recorded in the document's XMP metadata with the current modification time of the actual file. A difference of **more than 5 seconds** marks the link as needing an update. If the file does not exist, the status becomes "⚠ Missing".

### PPI and scale

PPI is calculated from the linked file's actual pixel dimensions and its placed size as `pixels × 72 ÷ placed size (pt)`, then averaged across width and height and rounded. This matches the Links panel's approach, but here the values line up in a table and can be sorted.

Pixel dimensions are read directly from the linked file's binary data. Supported formats are **PNG / JPEG / PSD**. For other formats (AI, PDF, TIFF, and so on), PPI and color space show `---`.

Scale is derived from the placed item's transformation matrix. Both axes are shown as `80.0% × 60.0%` only when they differ by 0.1% or more; otherwise a single value is shown.

### Color space

The color mode read from the linked file is shown together with the embedded ICC profile name, as in `RGB (sRGB IEC61966-2.1)`. Without a profile, only the color mode is shown. This is also limited to PNG / JPEG / PSD.

## Grouping identical files

By default, **rows for the same linked file are merged into one**. Files are treated as identical when their paths match. The Uses column shows the count, and the Artboard column lists every artboard the file is placed on (for example `1, 3, 5`).

Clearing the checkbox switches to one row per placement.

Note that **missing links cannot be grouped**. Their paths are unavailable, so identically named files still appear as separate rows.

The initial sort order is Uses, descending, so heavily reused images come first.

## Sorting

Sorting is controlled by the order dropdown and the ascending/descending radio buttons in the Sort panel.

| Order | Notes |
| --- | --- |
| File name | |
| Size | Only while the Size column is visible |
| Uses | Only while the Uses column is visible |
| Artboard | Only while no artboard filter is applied |
| Width / Height / Scale / PPI | Only while the size, %, and PPI columns are visible |
| Status | Missing → Needs update → Linked → Embedded |
| Color space | Only while the Color space column is visible |

**The dropdown entries follow the visible columns.** Hidden columns cannot be sorted on. Showing or hiding columns preserves the selected order whenever possible.

Choosing Size, Uses, Width, Height, Scale, or PPI **switches to descending automatically**, since the largest or most frequent entries are usually the ones of interest. Choosing File name or Status leaves the direction unchanged.

Rows with no value (such as a PPI of `---`) always collect at the end, in both ascending and descending order.

## Filtering

### Status

Four checkboxes — "✓ Linked", "⚠ Missing", "⟳ Needs update", and "▣ Embedded" — control which statuses are listed. To see only broken links, clear the other three. All of them are on by default.

### Artboard

Selecting an artboard in the dropdown narrows the list to images on that artboard.

**Artboards with no images are dimmed and cannot be selected**, so it is clear from the moment the palette opens which artboards are worth looking at.

Selecting an artboard also **fits the canvas view to that artboard**, keeping the list filter and the screen in sync.

The adjacent ◀ ▶ buttons step through artboards. Dimmed artboards are skipped automatically, and stepping past the last artboard wraps to the first (it does not return to "All"). The buttons are disabled entirely when one or fewer artboards are selectable.

Note that **the Artboard column disappears while an artboard filter is active**, since every row would show the same number.

An image is assigned to the artboard whose rectangle contains the center of the image's visible bounds. An image inside none of them shows `-`.

## Selection sync

Selecting a row also selects the corresponding placed image on the canvas and zooms and centers it to fill the view. For a row that groups identical files, every placement is selected together.

Clear the "Zoom on selection" checkbox to keep the view where it is.

## Displaying file paths

The path of the selected row appears in the File path panel. Three checkboxes control how it is shown.

| Checkbox | Behavior |
| --- | --- |
| Full path | Shows the absolute path as is |
| Shorten Dropbox path | Omits the Dropbox mount path prefix |
| File name | On shows the file name; off stops at the folder |

By default "Full path" and "File name" are off, so **the folder path with the home directory replaced by `~`** is shown. Paths tend to run long, so the short form is the default.

While "Shorten Dropbox path" is on, "Full path" is dimmed, since shortening and a full path are mutually exclusive.

Dropbox shortening uses the `DROPBOX_PREFIX` value at the top of the script, which should be set to your local mount path. Setting it to `""` **hides the checkbox entirely**.

## Linked folder list

Every folder referenced by the placed images is listed once, with duplicates removed. Selecting a row in the image list automatically highlights the folder containing that file.

Double-clicking, or clicking Open, reveals the folder in the Finder.

## Per-image operations

These buttons are available while a row is selected in the list.

- **Copy file name**: copies the file name to the clipboard
- **Open**: opens the linked file in its associated application
- **Rename**: renames the file on disk and relinks
- **Delete**: removes the placed image from the document
- **Relink / Relink all**: replaces the link target
- **Embed**: converts a linked image into an embedded image
- **Unembed**: turns an embedded image back into a link

While an embedded image is selected, everything that depends on a source file (Open, Rename, Relink, Embed) is disabled and only Unembed is available. Delete and Copy file name work for either kind of row.

### Rename

**The file on disk is renamed and relinked in one step**, collapsing the usual "rename the file, then fix the broken link" sequence into a single operation.

The extension is preserved automatically. The input field contains only the name without its extension, and an extension typed in by mistake is stripped. If a file with the new name already exists, an overwrite confirmation appears.

For a row that groups identical files, every placement is relinked to the new name together.

### Delete

When the target sits **inside a clipping group (mask)**, a dedicated dialog offers two ways to delete it.

| Option | Behavior |
| --- | --- |
| Delete the image only | Removes the placed image and keeps the clipping group |
| Delete the whole clipping group | Removes the clipping group surrounding the image |

For an image outside a clipping group, only a standard confirmation alert appears.

### Relink

The button label changes with the situation. When identical files are grouped and the row covers multiple placements, the button reads **Relink all**, and clicking it replaces every placement with the chosen file after a confirmation. With a single placement, it reads Relink.

### Embed

Converts the selected linked image into an embedded image. For a row that groups identical files, every placement is embedded together after a confirmation.

**Embedding a PSD turns its layers into a group**, so the group is ungrouped as well — but only when the image is not inside a clipping group, to leave the mask structure intact.

### Unembed

Turns the selected embedded image back into a linked image. A temporary action (`adobe_placeDocument`) is generated on the fly and played with "replace selection", so **Illustrator itself preserves the position, size, rotation, and stacking order**.

The link target is resolved in this order:

1. If the embedded image still knows its source file and that file exists, it links to it
2. Otherwise the image is **reset to 100% scale and 0° rotation and exported as a PSD into the `Links` folder**, and linked to that

The export name is looked up in the layer name, then the parent group name with an extension, then the original file name in the XMP manifest, then falls back to `image1`, `image2`, and so on. A numbered suffix is added when a file of the same name exists, so nothing is overwritten.

**Collect after relinking** (on by default) copies the linked file into the `Links` folder next to the document and repoints the link. A numbered suffix is added for different files of the same name, identical files are not copied again, and an image that was just exported as a PSD is already in the folder so it is not copied twice.

Locked or hidden layers, groups, and images cannot be selected and therefore cannot be replaced. Unsupported color spaces (anything other than CMYK / RGB / Grayscale) are skipped, and every reason is collected into a single alert.

Because a PSD export may be involved, an unsaved document fails: there is no place to create the `Links` folder.

## Folder-level operations

These are used after selecting a row in the linked folder list.

### Relink folder

Images referencing the selected folder are **replaced in bulk with identically named files in another folder**. Files with no match in the destination folder are skipped, and the counts appear in the status line.

### Change extension

Choosing a reference folder and an extension relinks each image to **a file with the same base name but a different extension**. This suits workflows such as swapping working PSDs for a delivery format.

The available extensions are png / jpg (jpeg) / psd / tiff (tif) / webp / avif / gif / ai / pdf. The default is psd.

`jpg / jpeg` and `tiff` **try two extensions in order**: `.jpg` (`.tif`) first, then `.jpeg` (`.tiff`) if that is missing. Extension case is ignored.

Images already referencing the same file are counted as skipped, and the totals for targets, successes, skips, and failures appear in the status line.

### Collect links

**A `Links` folder is created next to the document, the linked files are copied into it, and the images are relinked** — the usual step when packaging data for handoff.

If a file with the same name already exists in the `Links` folder, it is not copied and the image is relinked to the existing file. Images already referencing the `Links` folder are skipped.

An unsaved document cannot run this, since there is no location to create the folder in.

## Other buttons

- **Open the Links panel**: opens Illustrator's standard Links panel
- **Refresh**: reloads the data from the document

The palette stays open, so click Refresh after replacing or re-placing images on the canvas. The previously selected row is reselected after reloading whenever possible.

## Target

PlacedItem objects (linked images placed in the document) and RasterItem objects (embedded images).

PPI and color space are only read from PNG / JPEG / PSD files. Embedded images have no readable source file, so they show no PPI.

## Notes

### Persistent palette and BridgeTalk

The script runs in a persistent engine via `#targetengine`. That engine's `app` loses its connection to the DOM while the palette is displayed, so every operation that touches the DOM — reading the selection, analyzing placements, moving the view, changing attributes — is collected into worker functions and delegated to the main engine (`target=illustrator`) through BridgeTalk on each action.

The worker functions are all registered in `WORKER_FUNCS` and installed into the main engine once on first use (`$.global.__LIM`); only the call expressions are sent afterwards. Sending every function body each time is noticeably slower given how many lines are involved. **Adding a worker function means registering it in `WORKER_FUNCS` as well** — a missing entry breaks only the delegating side, silently.

Worker function bodies are concatenated with `toString()`, which imposes a few rules:

- No line comments (`//`) — the newline is lost, which would comment out the code that follows. Block comments only.
- Always terminate statements with a semicolon.
- Cross-references use the `w_` prefix names (everything lives in the same IIFE closure).

Return values use a marker scheme: `OK\n<json>` / `NODOC` / `ERR\n<msg>`, distinguished by the first line.

### How itemIndex is stored

Each row carries an `itemIndex`: an index into `placedItems` for a linked image, or into `rasterItems` for an embedded one. Those two index spaces collide, so **embedded images add 1000000** to theirs. On the worker side, `w_itemByIndex()` reads that value to decide which collection to look in. The palette side keeps treating it as a plain integer, so the selection-restore and matching code stays unchanged.

### Unembed and the bridge timeout

BridgeTalk sends wait 10 seconds by default, which is not always enough for unembedding: it may include a PSD export and an action playback. That one call waits up to 300 seconds and **is never resent on timeout** — resending a write operation would place the image twice.

### Handled on the palette side

`File` operations such as existence checks, renaming, and copying do work in the persistent engine, so those stay on the palette side. For relinking, only the link replacement itself is delegated.

Embedding and unembedding are mostly DOM work, so each one runs as a single worker call. The PSD export behind unembedding (creating a temporary document and calling `exportFile`) and the generation and playback of the temporary action also happen on the worker side.

### Other

Copying a file name to the clipboard works by creating a temporary text frame and calling `app.copy()`. The original selection is restored afterwards.

The ◀ ▶ artboard stepper draws its triangles as vectors on the button face rather than relying on font glyphs, and picks its color from the UI brightness setting.

## Article

[Redesigning the Links panel with a script (Japanese)](https://note.com/dtp_tranist/n/na66732d2056a)

## Changelog

- v1.5.1 (2026-07-27): Fixed the file extension being lost from the file name when embedding
- v1.5.0 (2026-07-27): Added embedded images to the list (a "▣ Embedded" status and filter) plus the Embed, Unembed, and Collect after relinking controls
- v1.4.2 (2026-07-27): Fixed URI-encoded file names in the list, rebuilt the artboard dropdown on Refresh, and guarded against re-entrant loads and double-clicked actions
- v1.0.0 (2026-04-24): Initial version
