# アートボードと同じ大きさの背景をテンプレート化して最背面へ

[![Direct](https://img.shields.io/badge/Direct%20Link-bg--template.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/bg-template.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/bg-template.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

現在のアートボードと同じ大きさの長方形を作成し、「bg-template」レイヤーに置いてテンプレート化したうえで最背面へ移動します.

### 主な機能

- CMYKドキュメントでは塗りをK45に設定（ダイアログで変更可）
- RGBドキュメントでは塗りを #999999 に設定
- 「bg-template」レイヤーを作成してテンプレート化し、最背面へ移動

### 使い方

1. 対象のアートボードをアクティブにします。
2. スクリプトを実行します。

### 注意点

- テンプレート化にはダイナミックアクションを使います。

### 更新履歴

- v1.0 (2025-07-29)
- v1.0.2 (2026-09-27): 数値欄の↑↓キーで RGB を変えたときも HEX を更新。↑↓以外のキーで値が丸め直される不具合を修正。項目名のコロンを言語別に
- v1.1.0 (2026-09-27): 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
