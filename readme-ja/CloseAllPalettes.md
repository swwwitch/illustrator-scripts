# 常駐パレットをまとめて閉じる

[![Direct](https://img.shields.io/badge/Direct%20Link-CloseAllPalettes.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/CloseAllPalettes.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloseAllPalettes.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

常駐エンジンで動いている各種フローティングパレットをまとめて閉じるユーティリティ。

- `$.global` はエンジンごとに独立しており、Illustrator では外から別エンジンへコードを届ける手段が無い（BridgeTalk 本文の `#targetengine` も `//@targetengine` も無視され、main エンジンで動く）
- そこで対象パレットと同じ共通エンジン `SwwwitchPalettes` で動き、`$.global.<参照名>` を直接読んで開いていれば `close()` する
- 共通エンジンへ移行していないパレットは閉じられない
- パレット参照は「Window を直接保持」する形式と「{ window: Window } のラッパー」形式の両方に対応する
- 閉じた後は `$.global.<参照名>` を null にして参照を解放する（各パレット本体の onClose でも解放されるが保険）
- 対象は PALETTES テーブルで管理

### 実行時の要点 / Runtime notes

- 共通エンジンでは IIFE の外の変数・関数がパレット同士で共有される。このスクリプトも基本情報ブロック以外は IIFE の中に置く

### 対象

AiMemoPalette / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManagerPalette /
UnifiedTypePalette / ImportAndApplyGraphicStylePalette / ArtboardDisplayPresetManagerPalette /
TextCountStatsPalette / SelectionInspectorPalette / ApplyLeadingPerTextFramePalette / TextProcessingPalette /
AiAlignToArtboardPalette / AiSmartRotateViewPalette / AutoKerningPalette / FontPresetPickerPalette / KPTSketchyPalette /
LockHistoryPalette / PathInspectorPalette / QuickTransformPalette / TypeBasicsPalette /
ArtboardNavigatorPalette / LEConvertToShapePalette / AiSmartPathfinderPalette / SmartDistributorPalette /
AiAdjustVerticalGapPalette / DirectPrefsPalette / DocumentFontListSelectorPalette / FavoriteFontPickerPalette

### スクリプト情報

- バージョン: v1.1.2

### 更新履歴

- v1.1.2（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.1.1（2026-10-03）ViewTogglePalette を対象に追加。
- v1.1.0（2026-10-03）どのパレットも閉じられなかった不具合を修正。BridgeTalk で送った本文の `#targetengine` は無視され main エンジンで動いていたため、パレットと同じ共通エンジン `SwwwitchPalettes` で直接閉じる方式に変更。対象の常駐パレット28本も共通エンジンへ移行した。
- v1.0.5（2026-10-02）FavoriteFontPickerPalette を対象に追加。
- v1.0.4（2026-09-26）TextFontPanelReinvented（削除）を対象から外した。常駐パレットのファイル名を末尾 Palette に改名したのに合わせて対象名を更新。
- v1.0.3（2026-09-26）開いているパレットが無いときのアラートを表示しないように変更。
- v1.0.2（2026-09-25）AiAdjustVerticalGap / DirectPrefs / DocumentFontListSelector / TextFontPanelReinvented を対象に追加。
- v1.0.1（2026-09-25）ArtboardDisplayPresetManager のエンジン名・参照名が本体と食い違っていて閉じられなかったのを修正。常駐パレット13本（AiAlignToArtboard ほか）を対象に追加。
