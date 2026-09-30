# 各行の幅を最長行にそろえて変倍

[![Direct](https://img.shields.io/badge/Direct%20Link-DynamicTextGeneratorSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/DynamicTextGeneratorSimple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGeneratorSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した1つのテキストフレームの各行をアウトライン幅で測定し、最長行の幅にそろうよう行ごとの文字サイズを変倍します。

### 主な機能

- ダイアログを開いたまま結果をプレビュー
- 行送りを自動（既定100%）に切り替えるかどうかを選択

### 使い方

1. テキストフレームを1つ選択します。
2. スクリプトを実行します。
3. プレビューを確認して［OK］をクリックします。

### 注意点

- パス上文字への変換まで行いたい場合は DynamicTextGenerator.jsx を使用してください。

### 更新履歴

- v1.1.5（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.1.2 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.1.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0.0 (2026-08-11)
