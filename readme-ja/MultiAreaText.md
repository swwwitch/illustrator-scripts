# 複数のテキストをエリア内文字にまとめる／分割する

[![Direct](https://img.shields.io/badge/Direct%20Link-MultiAreaText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/MultiAreaText.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MultiAreaText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

複数のテキストを1つのエリア内文字にまとめたり、逆に分割したりします。

### 主な機能

- マージ: テキストを2つ以上選択して、内容を統合した新しいエリア内文字を作成

### 使い方

1. 対象のテキストを選択します。
2. スクリプトを実行します。

### 注意点

- タブや段落を手がかりに分割・連結したい場合は TextProcessingPalette.jsx を使用してください。

### 更新履歴

- v1.2.3 (2026-09-29)：ダイアログの不透明度を98%に変更
- v1.2.2 (2026-09-28)：一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした。ボタン行を共通の部品で組むようにした
- v1.2.1 (2026-09-28)：ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.2.0 (2026-09-27)：数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.0 (2026-09-27)：グループ内のテキストにも対応。スレッドでないエリア内文字1つでは［エリア内文字の高さ］だけを実行。UIの文言を見直し。数値欄で↑↓キーによる増減に対応
- v1.0 (2026-03-04)
