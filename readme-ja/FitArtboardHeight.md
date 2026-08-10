# FitArtboardHeight

[![Direct](https://img.shields.io/badge/Direct%20Link-FitArtboardHeight.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/FitArtboardHeight.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardHeight.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトのバウンディングボックスに **上下マージン** を加味して、作業中のアートボードの **高さのみ** を自動調整（左右＝幅は固定）。
- **選択がない場合**は、各アートボード内のオブジェクトを対象にして、**すべてのアートボード**の高さを個別に調整。

### 主な機能

- ライブプレビュー（軽量化のため `visibleBounds` 固定）
- プレビューの再描画をデバウンス（軽量化）
- プレビュー境界の切替（確定時のみ `visible`/`geometric` を反映）
- 〈対象〉パネル：
  ・「作業アートボードのみ」＝選択ありはアクティブABのみ／選択なしは全AB（従来挙動）
  ・「すべてのアートボード」＝選択を無視して全AB
- 文字オブジェクトは一時アウトライン化して正確な境界を計測（処理後に復元）
- 単位ラベル表示、ダイアログ位置の記憶（バージョン別キー）

### 処理の流れ

1) ダイアログで上下マージンを指定
2) プレビューでアートボードの上下のみ更新（左右は保持）
3) OKで確定、Cancelで元に戻す

### 注意点

- プレビューは高速化のため `visibleBounds` 固定です。確定時に設定の `visible/geometric` が反映されます。

### 更新履歴

- v1.2 (20250825) : 〈対象〉パネル追加（全AB/作業AB 切替）。プレビューを `visibleBounds` 固定＋`app.redraw()` をデバウンス。単位ラベルとローカライズ整理。ダイアログ位置の保存キーをバージョン別に。
- v1.1 (20250825) : アートボード **高さのみ** 調整に簡素化。未定義 `doc` 参照を修正。コメント整理と説明文更新。
- v1.0 (20250825) : 初期バージョン

---

### スクリプト情報

- バージョン: v1.2
