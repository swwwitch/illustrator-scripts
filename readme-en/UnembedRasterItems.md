# UnembedRasterItems

[![Direct](https://img.shields.io/badge/Direct%20Link-UnembedRasterItems.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/UnembedRasterItems.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that replaces every embedded raster image in the document with a linked image.
- A dialog lets you choose between relinking to the original file and forcing a PSD export.
- When relinking, the pre-embed file is used directly if it still exists; otherwise a PSD is exported.
- Files are written to a "Links" folder next to the Illustrator document (created automatically when missing).
- Exported files reuse the pre-embed original file name, looked up in the layer name, the parent group name and the XMP manifest; when unavailable, image1, image2… is used instead.
- An incremental suffix (e.g. `myPhoto(2).psd`) is added when a file of the same name exists, so existing files are never overwritten.
- The image is exported at its placed size, and the linked item keeps the rotation, position and stacking order of the original.
- The pixel dimensions (resolution) of the original image are preserved.

### Main Features

- Processes every embedded raster image in the document at once (no selection needed)
- Method chosen in a dialog
  - **Relink to the original file**: follows the path in the XMP manifest and links to it when the file exists; falls back to exporting a PSD when the size does not match
  - **Force unembed**: always exports a PSD and replaces the item
- Original file name recovery, looked up in this order
  1. Layer name (when the item has been named)
  2. Parent group name that carries a file extension
  3. XMP manifest (the pre-embed file path recorded in `xmpMM:Manifest`)
  4. `image1`, `image2`… as a fallback
  - Strips the extension and replaces characters that are illegal in file names (`\ / : * ? " < > |`) with underscores
- Supports CMYK / RGB / Grayscale color spaces (others are skipped and reported)
- Detects images whose size changed because of an effect, and either warns or leaves them embedded
- Reports the result (converted count, warnings, errors) in a single alert
- Automatic Japanese / English UI

### Workflow

1. Read the current scale and rotation from the transformation matrix of the embedded image
2. Duplicate it into a temporary document and undo the rotation only, **keeping the placement scale**
3. Fit the artboard to the image and export as PSD at the effective resolution derived from the scale (placed at 24% = 7200 / 24 = 300 ppi, which keeps the original pixel dimensions)
4. Place the exported PSD as a linked item and restore the rotation, position and stacking order (no scaling needed, since the PSD is already at the placed size)
5. Delete the original embedded image

### User Settings

Edit the "User settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `LINKS_FOLDER_NAME` | `"Links"` | Name of the export folder |
| `USE_XMP_NAMES` | `true` | Use the original file names from the XMP manifest when the layer name is empty |
| `PRESERVE_RESOLUTION` | `true` | Derive the effective resolution from the placement scale, preserving the original pixel dimensions |
| `EXPORT_RESOLUTION` | `300` | Fixed resolution (ppi) used when `PRESERVE_RESOLUTION` is `false`; resamples the image |
| `SIZE_TOLERANCE` | `0.01` | Tolerance for treating two sizes as equal (pt) |
| `SKIP_RESIZED_ITEMS` | `false` | When `true`, items whose size does not match are left embedded |
| `KEEP_EXPORT_DOC_OPEN` | `false` | When `true`, the temporary export document stays open (for testing) |

### Notes

- Unembedding raster items with effects (e.g. drop shadow) applied is not recommended.
  - The effect is rasterized into the linked image
  - Its resolution follows the embedded image, not the document raster settings
  - Some effects shift the position of the linked image compared to the embedded one
- Setting `PRESERVE_RESOLUTION` to `false` resamples the image.
- Images whose effective resolution exceeds 2400 ppi are exported at 2400 ppi, with a warning in the result.
- Opacity, blending mode and other appearance settings on the embedded item are not carried over to the linked item.
- Manifest names are matched by order **only when the de-duplicated count equals the number of embedded images**. The manifest can keep entries from earlier embeds, and when the counts disagree the script falls back to `image1`…

### Differences from the Original

- Uses the pre-embed original file name for the exported file, from the layer name, the parent group name or the XMP manifest; the original script always used image1, image2…
- Skips unsupported color spaces, failed exports and items whose scale cannot be determined, and reports them together
- Checks for no open document, an unsaved document and documents without embedded images before running
- Japanese / English messages
- The handling of items resized by effects moved from an `if (false)` branch to the `SKIP_RESIZED_ITEMS` user setting
- Export folder name, resolution and other options are collected as user settings at the top
- Split into smaller functions, each documented with JSDoc

### Not Supported

- No open document (an alert is shown and the script exits)
- Unsaved document (the export folder cannot be resolved; an alert is shown and the script exits)
- Documents without embedded raster images (an alert is shown and the script exits)
- Embedded images whose color space is not CMYK, RGB or Grayscale (skipped and reported)
- Linked images (PlacedItem)

### Credits

- Original script by m1b
- [Adobe Community: Is it possible to convert rasterItem to placedItem?](https://community.adobe.com/t5/illustrator-discussions/is-it-possible-to-convert-rasteritem-to-placeditem/m-p/13081172)

### Update History

- v1.4.0 (20260727): Current version
