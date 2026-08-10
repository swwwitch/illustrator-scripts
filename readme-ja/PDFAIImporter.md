# PDFAIImporter

[![Direct](https://img.shields.io/badge/Direct%20Link-PDFAIImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/PDFAIImporter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAIImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、現在のドキュメント上にページを配置します。
配置方法は「各ページを個別のアートボードとして並べる」「アートボードを追加せずオブジェクトとして配置する」から選べます。

### 対象

- 全ページ（初期値）
- 先頭ページのみ
- 指定ページ（例: 1-10, 1,3,5）

読み込みファイルが確定するまでは、対象ページパネルと配置方法パネルは無効です。
選択中の配置画像、または［ファイル指定］で選んだ PDF/AI ファイルから総ページ数を推定し、
対象ページパネルに「実際に配置されるページ数／総ページ数」を表示します。
［ファイル指定］のダイアログでは PDF/AI ファイルのみを選択対象にします。

### 配置方法

- アートボードごと：各ページサイズに合わせてアートボードを作成／更新し、倍率 100% 固定で配置
- アートボードを無視：現在のアートボード左上を基準に、指定倍率でオブジェクトとして配置（配置後は対象を選択し、全体が見えるよう表示を調整）

### レイアウト

- 列数を指定すると、その数ごとに改行してグリッド状に配置
- 列数が「自動」のときは、カンバス右端に達したタイミングで折り返し
- 列数指定時のみ、列数入力欄の右側に推定の行 × 列を表示
- 間隔は配置方法で意味が変わる（アートボードごと＝アートボードの間隔／無視＝オブジェクト同士の間隔）

### ケイ

- なし
- ケイ線のみを追加（角丸オプションを有効にすると角丸半径を指定可能）

### 操作

- 列数・間隔・倍率の入力欄は ↑↓ で ±1、Shift+↑↓ で ±10
- 列数が「自動」のときは ↑ で 1 に切り替え
- 配置完了後は結果全体が見えるよう表示を自動調整（アートボードごと＝全アートボード表示／無視＝配置オブジェクト全体表示）

### 紹介記事

https://note.com/dtp_tranist/n/n42595650216f

### スクリプト情報

- バージョン: v1.1.1
- 最終更新: 2026-04-13
