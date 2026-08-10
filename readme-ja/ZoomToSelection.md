# ZoomToSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-ZoomToSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ZoomToSelection.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZoomToSelection.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトに合わせて、アクティブビューをズーム＆センタリングします。
- 複数選択時は全体のバウンディングボックスにフィットし、周囲に少し余白を残します。
- 選択が無い場合は、現在位置のまま 100% 表示に戻します。
- ズームは一気に切り替えず、少しずつ補間してアニメーション風に動かします。

### 設定

- `ZOOM_FIT_RATIO`：フィット時に残す余白の係数（1.0 でぴったり、0.9 で 10% の余白）
- `MAX_ANIMATION_STEP_COUNT`：補間ステップ数の上限
- `FRAME_DELAY_MS`：各フレームの待機ミリ秒

### オリジナルアイデア

John Wundes - Zoom and Center to Selection v2.

http://www.wundes.com/js4ai/copyright.txt

アニメーション補間は ArtboardNavigator.jsx（古島佑起さん）を参考にしています。
