# ApplyLeadingPerTextFrame

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyLeadingPerTextFrame.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ApplyLeadingPerTextFrame.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyLeadingPerTextFrame.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択された各テキストフレームの各行について、行頭数文字のフォントサイズを基準に行送りを再計算して適用します。

適用する行送りの割合はダイアログで指定できます。

### 使い方

1. 対象のテキストフレームを選択します。
2. スクリプトを実行します。
3. 行送りの割合を指定して実行します。

### 注意点

- テキストの一部（TextRange）を選択している場合は、その親のテキストフレームに正規化して処理します。
- 行送りの値そのものではなく、自動行送りの値（％）を変更して調整します。
- 割合を固定した派生版として ApplyLeadingPerTextFrame110.jsx / 150.jsx / AUTO.jsx があります。

### 更新履歴

- v1.1.0 (2026-07-08)
