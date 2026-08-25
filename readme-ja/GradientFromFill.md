# GradientFromFill

[![Direct](https://img.shields.io/badge/Direct%20Link-GradientFromFill.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/GradientFromFill.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GradientFromFill.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した塗りオブジェクトに対して、元の塗り色を始点にした線形グラデーションを作成します。

### 主な機能

- 単色オブジェクトはその塗り色を始点に使用
- 塗りがグラデーションの単一オブジェクトでは、黒・白・透明を除いたストップから始点色を選択
- 終点カラーは黒／白／透明／補色／淡色から選択
- 角度は 0 / 30 / 45 / 60 / 90 度から選択
- セパレートグラデーション、反転、プレビューに対応

### 使い方

1. 対象のオブジェクトを選択します。
2. スクリプトを実行します。
3. 始点カラー・終点カラー・角度を指定して［OK］をクリックします。

### 注意点

- 複数オブジェクトを選択している場合は、始点カラーパネル全体が無効になります。
- キャンセル時は元の塗り色へ戻します。プレビュー中に例外が発生した場合も、元の塗り色と選択状態を復元します。
- 複合パス、グループ内の再帰処理、クリッピンググループ内のオブジェクトにも対応します。

### 更新履歴

- v1.1.0
