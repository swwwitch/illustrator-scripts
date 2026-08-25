# PDFAISpreadImporter

[![Direct](https://img.shields.io/badge/Direct%20Link-PDFAISpreadImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/PDFAISpreadImporter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAISpreadImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、新規ドキュメント上に各ページを個別のアートボードとして配置します。

横長ページは見開きとして自動判定し、左右2つのアートボードに分割して配置します。

### 主な機能

- ページ範囲の指定（全ページ／先頭ページのみ／指定ページ）
- 横長ページを見開きとして自動判定し、左右に分割
- 偶数ページの位置を右／左から選択
- PDFのトリミング設定（アート／トリミング／仕上がり／裁ち落とし）を選択
- 新規ドキュメントのカラーモードを CMYK / RGB から選択

### 使い方

1. スクリプトを実行します。
2. 読み込む PDF/AI ファイルを指定します（選択中の配置画像があればそれを使います）。
3. ページ範囲、偶数ページの位置、トリミング設定、カラーモードを指定します。
4. 実行すると、新規ドキュメントにアートボードが並びます。

### 注意点

- 選択中の配置画像、または指定ファイルからページ数を推定し、ページ範囲に反映します。
- アートボード間隔は 100 pt 固定です。
- ラスタライズ効果解像度は 300 ppi 固定です。
- 各ページを分割せずに配置したい場合は PDFAIImporter.jsx を使用してください。

### 更新履歴

- v1.1.0 (2026-03-18)
