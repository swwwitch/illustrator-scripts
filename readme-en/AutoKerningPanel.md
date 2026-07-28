# Kerning Settings Palette

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoKerningPanel.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoKerningPanel.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerningPanel.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Auto kerning lives in the Character panel and proportional metrics lives in the OpenType panel. Both are settings you touch constantly in Japanese typesetting, yet they require keeping two panels open and moving between them.

This script puts just those two into a single persistent palette. It is the kerning part of the [Unified Type Panel](UnifiedTypePanel.md), pulled out on its own — lighter to reach for when kerning is all you need to adjust.

## Usage

1. Select the target text (a text frame, text nested in a group, or a range selected in text-edit mode).
2. Run the script.

**There is no Apply button.** The setting is applied to the selected text as soon as you operate a radio button or the checkbox. Press `Esc` to close the palette. Running the script again while the palette is already open simply brings it to the front instead of opening a second one.

Every time the palette regains focus it re-reads the current values of the selected text into the UI, so selecting different text updates the display along with it.

## Auto kerning

| Choice | Internal value |
| --- | --- |
| Metrics - Roman Only | METRICSROMANONLY (metrics for Roman only) |
| 0 | NOAUTOKERN |
| Metrics | AUTO |
| Optical | OPTICAL |

## Proportional metrics

Choosing "Metrics" turns proportional metrics on automatically. Any other choice turns it back off.

The checkbox can also be toggled on its own, which changes only proportional metrics and leaves the kerning method alone.

## Scope

**Settings apply to every paragraph the selection touches.** Selecting only a few characters in text-edit mode still applies the change to the whole paragraph they belong to.

Kerning and proportional metrics are settings that cause trouble when they vary within a paragraph, so this is deliberate.

## About the values shown

The palette reads the **first readable paragraph's first character** in the selection. With a multi-frame selection, or when settings are mixed within one frame, the display shows that first value and may not reflect the whole selection.

## Targets

Text frames, text nested in groups (processed recursively), and TextRange selections in text-edit mode.

## Notes

A persistent Illustrator palette loses its connection to the document DOM while it is displayed, so every actual object operation is delegated to the main engine over BridgeTalk. The delegated code is wrapped with `encodeURIComponent` before it is sent.

The palette position is saved to `Folder.userData` on every move and restored on the next launch. A saved position that would land off-screen after a monitor layout change is ignored, and the palette opens at the default spot.

### note

- [自動カーニングとプロポーショナルメトリクスのパレット｜DTP Transit 別館](https://note.com/dtp_tranist/n/ne7a198a4f527)
