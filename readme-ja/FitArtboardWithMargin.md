# FitArtboardWithMargin.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-FitArtboardWithMargin.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/FitArtboardWithMargin.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトの外接範囲にマージンを加えてアートボードサイズを自動調整
- 単位に応じたマージン初期値を設定し、即時プレビュー可能なダイアログを提供

![](https://www.dtp-transit.jp/images/ss-406-318-72-20250709-065622.png)

### 主な機能

- 定規単位ごとのマージン初期値設定
- 外接バウンディングボックス計算
- 即時プレビュー付きダイアログ
- ピクセル整数値への丸め

### 処理の流れ

1. 選択オブジェクトがない場合は全オブジェクトを対象
2. 単位に応じて初期マージン値を設定
3. ダイアログでマージンを入力（即時プレビュー）
4. バウンディングボックスとマージン適用
5. アートボードを更新

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- 数字入力支援（+/-ボタン、0ボタン）
- ダイアログボックスを閉じずにプレビュー更新
- 単位系（mm、px など）によってデフォルト値を切り替え
- アートボードの座標・サイズをピクセルベースで整数値に
- オブジェクトを選択していない場合には、すべてのオブジェクトを対象に

### note

https://note.com/dtp_tranist/n/n15d3c6c5a1e5

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.1 (20250708) : UI改善、ポイント初期値変更
- v1.2 (20250709) : UI改善とバグ修正