# Close every floating palette at once

[![Direct](https://img.shields.io/badge/Direct%20Link-CloseAllPalettes.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/CloseAllPalettes.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloseAllPalettes.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A utility that closes every floating palette running in a persistent engine.

- Each palette lives in its own persistent engine via `#targetengine`, and `$.global` is independent per engine, so one script cannot reach another engine's palette reference directly
- Instead, one BridgeTalk message carrying `#targetengine` is sent per engine; the receiving engine reads `$.global.<reference>` and calls `close()` when the palette is open
- Both palette reference shapes are supported: a `Window` held directly, and a `{ window: Window }` wrapper
- After closing, `$.global.<reference>` is set to null to release the reference (each palette's own `onClose` does this too, as a safety net)
- Targets are listed in the `PALETTES` table; adding a palette is a one-line change

### Runtime notes

- A BridgeTalk cross-engine round trip works synchronously even when run from File > Scripts. The sending script has to stay alive until the reply arrives, so `BridgeTalk.pump()` waits for it after each send (without the wait the script ends before delivery and the palette is never closed)
- Sent BridgeTalk objects are kept in an array until their reply arrives; if they are collected in the meantime the message is never delivered
- Palettes are closed one at a time so that `close()` calls from several engines do not overlap on the same UI thread and hang

### Target

AiMemoPalette / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManagerPalette /
UnifiedTypePalette / ImportAndApplyGraphicStylePalette / ArtboardDisplayPresetManagerPalette /
TextCountStatsPalette / SelectionInspectorPalette / ApplyLeadingPerTextFramePalette / TextBreakSplitMergePalette /
AiAlignToArtboardPalette / AiSmartRotateViewPalette / AutoKerningPalette / FontPresetPickerPalette / KPTSketchyPalette /
LockHistoryPalette / PathInspectorPalette / QuickTransformPalette / TypeBasicsPalette /
ArtboardNavigatorPalette / LEConvertToShapePalette / AiSmartPathfinderPalette / SmartDistributorPalette /
AiAdjustVerticalGapPalette / DirectPrefsPalette / DocumentFontListSelectorPalette

### Script info

- Version: v1.0.4

### Update History

- v1.0.4 (2026-09-26) Removed TextFontPanelReinvented (deleted) from the targets. Updated target names after renaming persistent palette scripts to end in "Palette".
- v1.0.3 (2026-09-26) No longer shows an alert when no palettes are open.
- v1.0.2 (2026-09-25) Added AiAdjustVerticalGap, DirectPrefs, DocumentFontListSelector and TextFontPanelReinvented.
- v1.0.1 (2026-09-25) Fixed ArtboardDisplayPresetManager not closing: its engine and reference names did not match the script. Added 13 more persistent palettes (AiAlignToArtboard and others).
