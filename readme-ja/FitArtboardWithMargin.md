# FitArtboardWithMargin.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-FitArtboardWithMargin.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/FitArtboardWithMargin.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトまたは全オブジェクトの外接バウンディングボックスにマージンを加え、アートボードを自動調整します。
- 定規単位に応じた初期マージン値設定と即時プレビュー可能なダイアログを提供します。
- ピクセル整数値に丸めたアートボードサイズを適用します。

<img alt="" src="https://www.dtp-transit.jp/images/ss-426-508-72-20250713-072422.png" width="50%" />

### 主な機能

- 定規単位ごとのマージン初期値設定
- 外接バウンディングボックス計算
- 即時プレビュー付きダイアログ
- ピクセル整数値への丸め

### 処理の流れ

1. 選択オブジェクトまたはアートボードを対象に選択
2. マージン値をダイアログで設定（即時プレビュー対応）
3. 設定値に基づきアートボードサイズを自動調整

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- ダイアログボックスを閉じずにプレビュー更新
- 単位系（mm、px など）によってデフォルト値を切り替え
- アートボードの座標・サイズをピクセルベースで整数値に
- オブジェクトを選択していない場合には、すべてのオブジェクトを対象に
- ↑↓キー、shift + ↑↓キーによる入力

### note

https://note.com/dtp_tranist/n/n15d3c6c5a1e5

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.1 (20250708) : UI改善、ポイント初期値変更
- v1.2 (20250709) : UI改善とバグ修正
- v1.4 (20250713) : 矢印キーによる値変更機能を追加、UI改善