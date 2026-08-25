# bg-template

[![Direct](https://img.shields.io/badge/Direct%20Link-bg--template.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/bg-template.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/bg-template.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

現在のアートボードと同じ大きさの長方形を作成し、「bg-template」レイヤーに置いてテンプレート化したうえで最背面へ移動します.

### 主な機能

- CMYKドキュメントでは塗りをK20に設定
- RGBドキュメントでは塗りを #999999 に設定
- 「bg-template」レイヤーを作成してテンプレート化し、最背面へ移動

### 使い方

1. 対象のアートボードをアクティブにします。
2. スクリプトを実行します。

### 注意点

- テンプレート化にはダイナミックアクションを使います。アクション定義は配列と `join("\n")` で組み立てる必要があり、現在の実装は `'''` を使っているため ExtendScript では構文エラーになります。

### 更新履歴

- v1.0 (2025-07-29)
