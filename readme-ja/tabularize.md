# tabularize

[![Direct](https://img.shields.io/badge/Direct%20Link-tabularize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/tabularize.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/tabularize.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトを「表」として解釈し、表組み用の塗りと線（横ケイ／縦ケイ）を生成します。

### 主な機能

- 塗り: 通常／ゼブラ／行方向に連結／ヘッダー行のみ
- オプション: ガター（定規単位）、1行目をヘッダー行に
- 線: 縦ケイ（なし／列間のみ／すべて）。ガター0のときは連結して描画
- プリセット: 代表的な組み合わせを一括適用

### 使い方

1. 表として扱いたいオブジェクトを選択します。
2. スクリプトを実行します。
3. 塗りと線のオプションを指定して実行します。

### 注意点

- 計算のためにテキストを複製・アウトライン化しますが、元のテキストは編集可能なまま残ります（非破壊）。
- ダイアログの値はセッション内で復元され、Illustratorを再起動するとリセットされます。

### 更新履歴

- v1.2.3
