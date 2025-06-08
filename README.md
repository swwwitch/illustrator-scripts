# Adobe Illustrator Scripts

This is a collection of JS scripts for Adobe Illustrator. 

## SmartSlice-with-Puzzlify

This script allows you to split selected images or vector artwork in Adobe Illustrator into grid or jigsaw puzzle shapes. Each piece can be offset and individually masked. The jigsaw shapes can include randomized connectivity for more natural puzzle designs.

Originally created by Jongware on 7-Oct-2010
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

![Uploading ss-2092-1082-72-20250608-151834.png…]()


### Features

- Supports image files, vector shapes, symbols, and groups.
- Supports both grid and traditional/random jigsaw puzzle shapes.
- Allows custom row/column input or piece count with automatic calculation.
- Optional offset effect with unit-based input.
- Optional scatter effect to displace pieces.
- Works on single selected objects or multiple selections (auto-grouped).
- UI localized in Japanese and English (auto-detected).

### Workflow

1. **Select an Object**
   - Valid types: RasterItem, PlacedItem, SymbolItem, PathItem, GroupItem, CompoundPathItem
   - You can also select multiple objects (will be grouped and converted to a symbol).

2. **Open the Script**
   - Launch from `File > Scripts > Other Script...` or drag it into Illustrator.

3. **Configure the Dialog**
   - **Total Pieces**: Approximate number of puzzle pieces.
   - **Rows/Columns**: Manual or auto-filled based on total pieces.
   - **Shape Type**: Choose from "Traditional", "Random", or "Grid".
   - **Offset**: Enable and input the offset value (e.g., -2pt).
   - **Scatter**: Enable to apply random scatter displacement (e.g., 30pt range).

4. **Run the Script**
   - Click the “Run” button. The object is split and each piece is masked.

### Notes

- **Offset**:
  - Negative values shrink the pieces; positive values enlarge them.
  - The unit label matches Illustrator’s current ruler units.
- **Scatter**:
  - Pieces are moved randomly within a specified range if enabled.

### Supported Objects

- Single selected object of types:
  - Placed image
  - Embedded image
  - Vector artwork (path, group, compound path)
  - Symbol instance

### Unsupported Scenarios

- No open document
- No selection
- More than one object selected without grouping (will auto-group)

## Change Log
- **v1.0.0** – Initial version
- **v1.0.1** – Symbol support
- **v1.0.2** – Vector artwork support
- **v1.0.3** – Grid shape support
- **v1.0.4** – Offset & unit label support
