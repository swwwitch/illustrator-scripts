# AiCreateArtboardGuides-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-AiCreateArtboardGuides--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/AiCreateArtboardGuides-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCreateArtboardGuides-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

アートボードを基準にガイドを整理・作成するツール。次の3系統をダイアログでまとめて設定できる。

- ルーラーガイドの変換：アートボード内に重なるルーラーガイドを検出し、アートボード基準の直線ガイドに引き直す（マスターでON/OFF）
- 中心ガイド：各アートボードの垂直・水平の中心にガイドを作成
- エッジガイド：各アートボードの上下左右にガイドを作成（マスターON/OFF、既定はOFF）
- ガイドが1本も無くても、中心・エッジの作成だけ実行可能
- 作成したガイドはすべて「_guide」レイヤーに集約（無ければ自動作成、ロック/非表示は解除して使用）

設定

- 「外側に延長」「外側へ延長」で、ガイドをアートボードの外側へ延ばす量を指定
- 「すべてのアートボード」は変換側（重なる全アートボード）と中心・エッジ側（全アートボード／アクティブのみ）で個別に指定
- 入力値は現在のルーラー単位として扱い、内部で pt に換算（単位は環境設定の rulerType を参照）

プレビュー

- 設定変更に追従するライブプレビュー（専用レイヤーに色付き線で仮表示し、確定時に本物のガイドへ置換）

### 値の変更

- ↑↓キーで±1増減
- shiftキーを併用すると±10増減

### 紹介記事（note）

https://note.com/dtp_tranist/n/n56d9c936a364

### スクリプト情報

- バージョン: v1.1.0
