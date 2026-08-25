# AiApplySwatchesToSelection-dialog

[![Direct](https://img.shields.io/badge/Direct%20Link-AiApplySwatchesToSelection--dialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/AiApplySwatchesToSelection-dialog.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiApplySwatchesToSelection-dialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- A modal dialog that applies swatches or predefined colors to selected objects or text.
- Choose apply unit (object / character / word / line / paragraph) and order (as-is / reverse / random / fully random) with radio buttons.
- "Random" shuffles the color list once and cycles through it; "Fully random" draws a color per target, so no sequence repeats.
- Live preview on every change. [OK] commits the result; [Cancel] (Esc) reverts it.
- The swatches selected when the dialog opens are captured as chips and used (by swatch name) for applying. Even a single selected swatch is used.
- When no swatches are selected, auto colors are used (CMYK CM/CY/MY generation, RGB defaults).
- "Per word" staggers colors so each line starts on a different color.
- The dialog runs in the main engine, so DOM work is executed directly (no BridgeTalk delegation).

### Script info

- Version: v1.8.0
