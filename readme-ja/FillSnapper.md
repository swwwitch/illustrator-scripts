# FillSnapper

[![Direct](https://img.shields.io/badge/Direct%20Link-FillSnapper.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/FillSnapper.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillSnapper.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択中のオブジェクトを「動かす対象」と「スナップ基準」に分類し、対象のバウンディングボックスを最寄りの基準線へ合わせます。

パスはアンカーポイントを直接変形するため、クリップグループ内の子パスにも対応します。

### 使い方

1. 動かしたいオブジェクトと、基準にする罫線をまとめて選択します。
2. スクリプトを実行します。
3. 許容差とスナップ距離を指定して実行します。

### オプション

**線判定の許容差**

細長いパスを水平線／垂直線として扱うための判定幅です。

**最大スナップ距離**

対象の辺からどれだけ離れた基準線まで吸着するかの制限です。0 の場合は距離制限なしになります。

### 更新履歴

- v1.0.1
