# AlignLabelWidths

[![Direct](https://img.shields.io/badge/Direct%20Link-AlignLabelWidths.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/_templates/AlignLabelWidths.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AlignLabelWidths.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ScriptUI の複数ラベル（statictext）の幅を、実際の描画幅を測って最長のものへ揃える再利用テンプレートです。

「ラベル：値」を縦に並べる情報パネルや設定パネルで、コロンの位置と値の開始位置をそろえる用途に使います。

### 主な機能

- 実際の描画幅を測ってラベル幅を揃えるため、ロケールや文言の長さに依存しません
- 固定文字数（`characters`）で指定する簡易版も用意しています
- ラベルと値を1行にまとめて追加するヘルパーを同梱しています

### 使い方

1. `alignLabelWidths()`（または簡易版 `setLabelsFixedWidth()`）と、必要なら `addLabelValueRow()` を対象スクリプトのトップレベルにコピーします。
2. 各行を作りながら、ラベルを配列に集めます。
3. 集めたラベル配列を揃え関数に渡します。
   - 方式A（推奨・実測して自動）: `alignLabelWidths(labels, 'right')`
   - 方式B（簡易・固定文字数）: `setLabelsFixedWidth(labels, 13, 'right')`

### 注意点

- `characters` の固定値より堅牢ですが、レイアウトが確定したあとに呼ぶ必要があります。
- テンプレートのため、単体で実行しても何も起こりません。

### 更新履歴

- v1.0.0: 初版（テンプレート）
