# ColorToK100Converter

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorToK100Converter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/ColorToK100Converter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorToK100Converter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2025-06-15
- RGB または CMYK で構成された黒を安定した K100 の黒に変換する Illustrator 用スクリプト
- テキスト、パス、スウォッチの塗りおよび線カラーが一括対象

### 主な機能

- RGB 黒（RGB 各値が 39 未満）を K100 に変換
- CMYK の多色ブラック（CMY 全てが 70 以上 または合計が 310 以上）を K100 に変換
- テキストの文字単位での変換対応
- スウォッチカラーも自動変換
- 日本語／英語インターフェース対応

### 処理の流れ

1) テキストの文字単位の塗り・線カラーを変換
2) パス、コンパウンドパスの塗り・線カラーを変換
3) グループ内オブジェクトを再帰的に変換
4) スウォッチのカラー定義を変換

### 更新履歴

- v1.0.0 (2025-06-12) : 初期バージョン
- v1.0.1 (2025-06-15) : 処理構造とコメント整理

### スクリプト情報

- バージョン: v1.0.1
