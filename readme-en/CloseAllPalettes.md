# Close every floating palette at once

[![Direct](https://img.shields.io/badge/Direct%20Link-CloseAllPalettes.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/CloseAllPalettes.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloseAllPalettes.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A utility that closes every floating palette running in a persistent engine.

- `$.global` is independent per engine, and Illustrator offers no way to run code in another engine from outside (a BridgeTalk body's `#targetengine` or `//@targetengine` is ignored and runs in the main engine)
- Instead, the script runs in the same shared engine as the palettes, `SwwwitchPalettes`, reads `$.global.<reference>` directly, and calls `close()` when the palette is open
- Palettes not yet moved to the shared engine cannot be closed
- Both palette reference shapes are supported: a `Window` held directly, and a `{ window: Window }` wrapper
- After closing, `$.global.<reference>` is set to null to release the reference (each palette's own `onClose` does this too, as a safety net)
- Targets are listed in the `PALETTES` table

### Runtime notes

- In the shared engine, variables and functions outside an IIFE are shared between palettes, so everything except the basic info block lives inside an IIFE

### Target

AiMemoPalette / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManagerPalette /
UnifiedTypePalette / ImportAndApplyGraphicStylePalette / ArtboardDisplayPresetManagerPalette /
TextCountStatsPalette / SelectionInspectorPalette / ApplyLeadingPerTextFramePalette / TextProcessingPalette /
AiAlignToArtboardPalette / AiSmartRotateViewPalette / AutoKerningPalette / FontPresetPickerPalette / KPTSketchyPalette /
LockHistoryPalette / PathInspectorPalette / QuickTransformPalette / TypeBasicsPalette /
ArtboardNavigatorPalette / LEConvertToShapePalette / AiSmartPathfinderPalette / SmartDistributorPalette /
AiAdjustVerticalGapPalette / DirectPrefsPalette / DocumentFontListSelectorPalette / FavoriteFontPickerPalette

### Script info

- Version: v1.1.2

### Update History

- v1.1.2 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.1.1 (2026-10-03) Added ViewTogglePalette to the targets.
- v1.1.0 (2026-10-03) Fixed palettes never closing: a BridgeTalk body's `#targetengine` was ignored and ran in the main engine. The script now closes palettes directly from the shared engine `SwwwitchPalettes`. All 28 target palettes have moved to that engine.
- v1.0.5 (2026-10-02) Added FavoriteFontPickerPalette to the targets.
- v1.0.4 (2026-09-26) Removed TextFontPanelReinvented (deleted) from the targets. Updated target names after renaming persistent palette scripts to end in "Palette".
- v1.0.3 (2026-09-26) No longer shows an alert when no palettes are open.
- v1.0.2 (2026-09-25) Added AiAdjustVerticalGap, DirectPrefs, DocumentFontListSelector and TextFontPanelReinvented.
- v1.0.1 (2026-09-25) Fixed ArtboardDisplayPresetManager not closing: its engine and reference names did not match the script. Added 13 more persistent palettes (AiAlignToArtboard and others).
