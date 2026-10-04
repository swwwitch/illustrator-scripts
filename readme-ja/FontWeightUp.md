# 選択テキストのフォントを一つ上／下のウェイトに切り替え

[![Direct](https://img.shields.io/badge/Direct%20Link-FontWeightUp.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontWeightUp.jsx)

[![Direct](https://img.shields.io/badge/Direct%20Link-FontWeightDown.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontWeightDown.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontWeightUp.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択しているテキストのフォントを調べ、同じファミリーの一つ上（FontWeightUp）または一つ下（FontWeightDown）のウェイトを適用する Illustrator スクリプト
- ウェイトを見本で見比べて選ぶときは [FontWeightPicker](FontWeightPicker.md) を使います

### 主な機能

- 同じファミリー・同じ系列（イタリック、Condensed などの字幅、Display などの用途）の中で一つ太い／細いスタイルに切り替え
- 文字ごとに異なるフォントが混在していても、それぞれ一段ずつ上げる／下げる
- ウェイトの判定は [TypefaceSampler](TypefaceSampler.md) と同じ
  - W3・W6、45 Light などの数値
  - L / R / M / DB / B / EB / H / U などの略号
  - Hair〜Ultra Black の英語のウェイト名（SemiBold・XBold なども可）

### 使い方

1. テキストフレームまたはグループを選択する（文字ツールで一部の文字を選択してもよい）
2. 太くするときは FontWeightUp、細くするときは FontWeightDown を実行する

### 注意点

- 別々のファミリーとして並ぶフォントは、スクリプト冒頭の `FAMILY_GROUPS` に細い順で登録すると一つのファミリーとして扱います（sw-L / sw-R / sw-B / sw-H は登録済み）
- 合成フォントと、ファミリー内で最も太い（Up）／最も細い（Down）ウェイトは変更しません。変更しなかったフォントは最後に一覧で表示します
- 1文字ずつ書き換えるため、長い文章では時間がかかることがあります

### 紹介記事

- [DTP Transit 別館｜note](https://note.com/dtp_tranist/n/n255437cfdba0)

### 更新履歴

- v1.0.0 (20260927) : 初期バージョン
- v1.0.1 (20260928) : sw-L / sw-R / sw-B / sw-H を一つのファミリーとして扱うように
- FontWeightDown v1.0.0 (20260928) : 一つ下のウェイトに切り替える FontWeightDown を追加
- v1.0.2（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

### スクリプト情報

- バージョン: FontWeightUp v1.0.1 / FontWeightDown v1.0.0
