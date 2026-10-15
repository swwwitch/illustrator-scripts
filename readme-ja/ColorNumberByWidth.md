# 全角・半角の英数字と等幅字形を色で塗り分け

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorNumberByWidth.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ColorNumberByWidth.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorNumberByWidth.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したテキストのうち、全角の英数字と「！？」を赤、半角の英数字と「!?」を青、等幅字形の文字を緑に塗ります。
- 全角・半角の混在や、等幅字形の指定を目で見て確かめるときに使います。

### 主な機能

- 対象の文字
  - 全角：０〜９、Ａ〜Ｚ、ａ〜ｚ、！、？
  - 半角：0〜9、A〜Z、a〜z、!、?
  - 等幅字形：OpenType の等幅全角字形・等幅半角字形・等幅1/3字形・等幅1/4字形を指定した文字（英数字以外も含む）
- 等幅字形は全角・半角より優先します。
- 色はダイアログボックスの色見本をクリックし、Illustrator 標準のカラーピッカーで指定できます。
- 色はドキュメントのカラーモード（RGB／CMYK）に合わせて作ります。

### 使い方

1. 対象のテキストを選択する（文字ツールで範囲を選択しても可）
2. スクリプトを実行する
3. 必要なら［全角］［半角］［等幅字形］の色見本をクリックして色を変える
4. ［OK］をクリックする

### 注意点

- 選択ツールで選んだときは、テキスト全体が対象です。グループの中のテキストも対象になります。
- 上記以外の記号（＃、％、# など）は塗りません。
- 色の初期値はスクリプト冒頭の `FULL_WIDTH_RGB` などで変えられます。選んだ色は次回には引き継がれません。

### スクリプト情報

- バージョン: v1.1.0

### 更新履歴

- v1.1.0（2026-10-16）等幅字形（等幅全角・半角・1/3・1/4字形）の文字を緑に塗る機能を追加。半角の色の初期値を #0099ff（CMYK は C100 M40）に変更
- v1.0.0（2026-10-16）初版
