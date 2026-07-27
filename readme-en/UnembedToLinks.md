# UnembedToLinks

[![Direct](https://img.shields.io/badge/Direct%20Link-UnembedToLinks.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/UnembedToLinks.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that replaces embedded raster images with linked images by generating and playing a temporary action (`adobe_placeDocument`).
- Because the action places the file with "replace selection" (rplc), Illustrator itself preserves the position, size, rotation and stacking order — the script does not restore them.
- The scope is chosen in a dialog: the current selection, the current artboard, or every embedded image in the document.
- Images whose original file is still known are linked to that file; when it is unknown, the embedded image itself is exported as a PSD and linked.
- With "collect after relinking" enabled, the linked file is copied into a "Links" folder next to the Illustrator document and the link is repointed to the copy.

### Main Features

- Scope chosen in a dialog
  - **Selected images only** (groups are searched recursively)
  - **Embedded images on the current artboard** (items overlapping the active artboard, tested with `geometricBounds`, so effects are not included)
  - **All embedded images**
- Target file list with two columns (file name / path)
  - **Full path**: shows the absolute path
  - **Shorten Dropbox path**: hides `DROPBOX_PREFIX`; when off, the home folder is shortened to `~`
- Images with an unknown original are always exported as a PSD into the "Links" folder and linked
  - The list shows the planned file name (`name.psd`) and "(export as PSD)"
  - The name is looked up in this order: layer name → parent group name that carries a file extension → XMP manifest → `image1`, `image2`…
  - The extension is stripped and characters that are illegal in file names (`\ / : * ? " < > |`) are replaced with underscores
  - An incremental suffix (e.g. `myPhoto(2).psd`) is added when a file of the same name exists, so existing files are never overwritten
- Collect after relinking (copy into the "Links" folder next to the document)
  - A `-1`, `-2`… suffix is added when a different file of the same name already exists
  - Files that can be considered identical (same path, or same size and modification date) are not copied again
- Reports the result (success / skipped / failed counts and details) in a single alert
  - When the relink succeeded but only the collect step failed, it is reported as a warning and still counted as a success
- The temporary action file (.aia) is always removed and the action set unloaded after playback

### Workflow

1. Collect the target images and resolve the export names for the whole document up front
2. Get the original file of each image (`item.file`); when unavailable, export a PSD as follows
   1. Duplicate the image into a temporary document
   2. Reset the scale to 100% and the rotation to 0°, then fit the artboard to the image
   3. Export a PSD into the "Links" folder and close the temporary document
3. Select only the target image, build the temporary `adobe_placeDocument` action, then `loadAction` → `doScript` → `unloadAction`
4. When collecting is enabled and the file is not a freshly exported PSD, copy it into the "Links" folder and repoint the link
5. Select the resulting linked images and show the summary

### Settings

Editable in the "Settings" block at the top of the script.

| Variable | Default | Description |
| --- | --- | --- |
| `ACTION_SET_NAME` | `"UnembedToLinksTempSet"` | Name of the temporary action set |
| `ACTION_NAME` | `"UnembedToLinksPlace"` | Name of the temporary action |
| `ACTION_FILE_NAME` | `"~/UnembedToLinksTemp.aia"` | Path of the temporary action file |
| `LINKS_FOLDER_NAME` | `"Links"` | Folder used for collecting and exporting |
| `EXPORT_RESOLUTION` | `72` | PSD export resolution (ppi). The image is reset to 100% before export, so the original pixel dimensions are kept |
| `USE_XMP_NAMES` | `true` | Use the original file names from the XMP manifest for images with no name |
| `DROPBOX_PREFIX` | Local Dropbox path | Prefix removed in the file list; set to an empty string to disable shortening |

### Notes

- When the document has not been saved, the "Links" folder cannot be located, so collecting and PSD export fail.
- Locked or hidden layers, groups and items cannot be replaced by the action and are reported as failures with the reason.
- Running the script repeatedly on the same document exports a new PSD each time for images with an unknown original (`image1.psd`, then `image1(2).psd`).
- PSD export supports CMYK / RGB / Grayscale only; other color spaces are reported as skipped.
- When the original is unknown and exactly one image is being processed, a manual file selection is offered if the export fails.
- Names from the XMP manifest are assigned in order **only when the number of unique candidates matches the number of unnamed images**; otherwise `image1`, `image2`… is used.
- Artboard membership is a rectangle overlap test, so an image spanning several artboards belongs to each of them.
- Effects such as drop shadows are rasterized into the exported PSD.

### Credits

- The PSD export logic is based on code by m1b.
- [Adobe Community: Is it possible to convert rasterItem to placedItem?](https://community.adobe.com/t5/illustrator-discussions/is-it-possible-to-convert-rasteritem-to-placeditem/m-p/13081172)

### Change Log

- v1.0.0 (20260727): Initial version
