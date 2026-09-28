# カラーパレットを生成してスウォッチに登録

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ColorGenerator.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorGenerator.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

カラーパレットを生成してアートボードに描画し、スウォッチグループにも登録します。

### 主な機能

- 対応アルゴリズム: Tailwind / Lightness / Saturation / Complementary / LCH / すべて
- 生成したパレットをアートボードへ描画
- 同じ内容をスウォッチグループとして登録

### 使い方

1. スクリプトを実行します。
2. アルゴリズムと基準色を指定します。
3. 実行すると、パレットが描画されスウォッチが登録されます。

### 更新履歴

- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.2 (20260928) : 英語表示で「HEX:」のコロンの後ろの余分な空白を削除（項目名のコロンを他のスクリプトとそろえた）
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0.1
