# Build a gradient from colors in layout order (blend version, merged into CreateGradientFromSelection)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

This script has been merged into [CreateGradientFromSelection.jsx](CreateGradientFromSelection.md).

To blend, turn on Blend duplicates in the CreateGradientFromSelection dialog.

### Update History

- 2026-09-29: Merged into CreateGradientFromSelection.jsx and removed
- v1.6
- v1.6.2: The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.6.3: Temporary actions now go through a shared load/play/unload routine, so the action set and temporary file are cleaned up even on failure
- v1.6.3: The button row is now built with the shared part
- v1.6.3: Clip groups are now measured by their mask (affects where the rectangle is placed)
