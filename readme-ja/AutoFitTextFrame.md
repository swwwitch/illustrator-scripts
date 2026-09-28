# エリア内文字に自動サイズ調整をワンクリックで適用

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoFitTextFrame.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoFitTextFrame.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoFitTextFrame.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したエリア内文字に［自動サイズ調整］を適用し、あふれた文字が収まるまでエリアの高さを広げます。
- [FitAreaText](FitAreaText.md) の［自動サイズ調整］だけを、ダイアログなしで実行する版です。

### 使い方

1. エリア内文字を選択します（複数可）。
2. スクリプトを実行します。

### 対象

- エリア内文字
- グループ内のエリア内文字（グループをたどって集めます）
- 文字カーソルで選択しているときは、そのエリア内文字

### 注意点

- 高さは広げるだけで、［自動サイズ調整］を適用したままにします（OFF には戻しません）。
- ポイント文字・パス上文字は対象外です。
- ロック・非表示・編集できないテキストは対象外です。
- ［自動サイズ調整］は一時アクションで適用します。途中で失敗したときは、残りのテキストを処理せずに警告を出します。

### 更新履歴

- v1.0.1 (2026-09-28) 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした。失敗したときは警告を出して止まるようにした
- v1.0.0 (2026-09-26) 初版
