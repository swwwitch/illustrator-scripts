# SmartBatchImporter.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartBatchImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/SmartBatchImporter.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A script to batch import multiple Illustrator files (.ai/.svg), paste only unlocked objects into a new document, and arrange them neatly.
- Supports adding filename labels, showing extensions, folder-based processing, progress bar, and various options.

### Main Features

- Batch import by folder selection
- Extract only unlocked objects
- Add filename labels (optionally show extensions)
- Choose to close or keep source files after import
- Color space, document size presets, and custom size
- Progress bar with cancel support
- Japanese and English UI support

### Process Flow

1. Select folder or open files
2. Configure color, size, label, and behavior options in dialog
3. Show progress bar and execute import and placement
4. Review result in a new document after completion

### Update History

- v1.0.0 (20250529): Initial version
- v1.0.1 (20250529): Changed folder import behavior, moved labels to "_label" layer
- v1.0.2 (20250529): Added file count display, progress bar, and cancel option
- v1.0.3 (20250529): Added progress count display (n/N)