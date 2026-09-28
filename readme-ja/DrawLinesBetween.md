# オブジェクトの間に水平の罫線を描く

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawLinesBetween.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/DrawLinesBetween.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawLinesBetween.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクト（図形／テキスト）を上から順に並べ、その間に水平の罫線を描画します。

### 主な機能

- 入力単位は環境設定の「線」（`strokeUnits`）に追従（表示ラベルと内部のpt換算の両方）
- ［延長］で罫線を左右方向に伸縮（＋で延長、−で短縮）

### 使い方

1. 罫線を挟みたいオブジェクトをまとめて選択します。
2. スクリプトを実行します。
3. 線幅や延長量を指定して実行します。

### 更新履歴

- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0
