# GroupEdgeAlignLEFT

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupEdgeAlignLEFT.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/GroupEdgeAlign-7/GroupEdgeAlignLEFT.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignLEFT.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択されているオブジェクト群の端または中心を取得し、左に揃えます。

揃え先は、アクティブアートボードの端、または条件に合うガイドです。整列方向はファイル名から自動判定されます。

### 主な機能

- `USE_GUIDES` が true のとき、指定方向に対応するガイドを探索し、`GUIDE_SEARCH_MODE` の条件に従って最適なガイドへスナップ
- 該当するガイドが無い場合はアートボード端に揃える
- `USE_GUIDES` が false のときは常にアートボード端に揃える

### 使い方

1. 揃えたいオブジェクトを選択します。
2. スクリプトを実行します。

### 注意点

- 中央揃え（左右中央／上下中央／上下左右中央）ではガイドを使わず、常にアートボード中央に揃えます。
- ダイアログで方向を選びたい場合は GroupEdgeAlign.jsx を使用してください。

### 更新履歴

- v1.0 (2025-04-06)
