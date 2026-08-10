# CollectArtboardTexts

[![Direct](https://img.shields.io/badge/Direct%20Link-CollectArtboardTexts.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/CollectArtboardTexts.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectArtboardTexts.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 全アートボード上にあるテキストフレームを収集し、最後のアートボードの右側に縦に並べて配置する
- Collect text frames on every artboard and place them as new text frames, stacked vertically to the right of the last artboard.

### 主な機能

- アートボードごとに重なるテキストフレームを抽出 / Pick text frames overlapping each artboard
- ロック・非表示のテキスト／レイヤー／グループは除外（親階層も走査）/ Skip locked or hidden frames and ancestor layers/groups
- 改行で結合せず、1 テキスト＝1 フレームで縦に並べる（スイッチで一括結合モードも可）/ One frame per source text stacked vertically (switch to combined mode available)
- 書き出した全フレームを選択してズーム表示 / Select the placed frames and zoom to fit

### 処理の流れ

1. アクティブドキュメントと出力先レイヤーを検証 / Validate the active document and target layer
2. アートボードごとに含まれるテキストを収集 / Collect contained texts per artboard
3. テキストごとに新規フレームを縦に並べて配置 / Place each text as a separate frame, stacked vertically
4. 書き出した全フレームを選択しズーム表示 / Select created frames and zoom in

### 更新履歴

- v1.0.0 (20260513) : 初版 / Initial release

### スクリプト情報

- バージョン: v1.0.0
