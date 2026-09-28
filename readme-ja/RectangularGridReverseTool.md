# 水平線・垂直線を長方形グリッドとして再構成

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangularGridReverseTool.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/RectangularGridReverseTool.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangularGridReverseTool.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した水平線・垂直線を解析し、長方形グリッドとして再構成します。

不揃いな罫線や、結合セルを含むレイアウトを整理し、整った格子構造に変換します。

### 主な機能

- 前処理：外枠を四辺に分割
- 配置：均等配置なし／均等（強制）／均等＋結合セル対応
- 対象：縦罫・横罫ごとに均等化の対象を制御
- 線（後処理）：突出線端、破線→実線、線幅（最大・最小・平均・指定）
- 後処理：外枠の長方形化、グループ化、ポイント文字をセル内で上下中央
- プレビュー：ダイアログを閉じずに結果を確認

### 使い方

1. 対象の水平線・垂直線を選択します。
2. スクリプトを実行します。
3. 前処理・配置・対象・線・後処理を指定し、プレビューを見ながら確定します。

### 更新履歴

- v1.3.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.3.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.2.0
