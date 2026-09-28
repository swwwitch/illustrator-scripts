# グラデーションをセパレートグラデーションに変換

[![Direct](https://img.shields.io/badge/Direct%20Link-Convert2SeparateGradient.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/Convert2SeparateGradient.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/Convert2SeparateGradient.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトのグラデーションを、指定した数のカラーストップでセパレートグラデーション（縞模様）に変換

### 主な機能

- 追加するカラーストップ数を指定（1 以上の整数）
- セパレートグラデーション（縞模様）への変換
- 特色（スポットカラー）の RGB / CMYK 自動変換
- RGB / CMYK ドキュメントに対応

### 既定の挙動

- 単一オブジェクトのみ対象
- 塗りがグラデーション以外、またはストップが 2 つ未満の場合はアラートで終了

### スクリプト情報

- バージョン: v1.1.4

### 更新履歴

- v1.1.4 (2026-09-29): ファイル名を `convert2separategradient.jsx` から `Convert2SeparateGradient.jsx` に変更
- v1.1.3 (2026-09-29): ダイアログの不透明度を98%に変更
- v1.1.2 (2026-09-28): ボタン行を共通の部品で組むようにした
- v1.1.1 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (2026-09-27): 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
