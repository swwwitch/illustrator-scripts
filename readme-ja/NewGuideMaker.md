# NewGuideMaker


[![Direct](https://img.shields.io/badge/Direct%20Link-NewGuideMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/NewGuideMaker.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

ダイアログで方向、位置、単位、対象（カンバス/アートボード）を指定してガイドを作成するスクリプト

<img alt="" src="https://www.dtp-transit.jp/images/ss-662-964-72-20250713-213514.png" width="70%" />

### 主な機能

- 水平方向・垂直方向のガイド作成
- 位置と裁ち落とし（マージン）の指定
- 単位選択（px, pt, mm）
- 環境設定に基づいた初期単位設定
- 上下キーによる数値調整

### 処理の流れ

1. ダイアログ表示
2. 設定入力
3. OK 押下でガイド作成

### メモ

Photoshopの［新規ガイド］ダイアログボックスをベースに、以下の要素を盛り込んでいます。

- アートボード、カンバスの選択
- アートボード選択時には裁ち落とし（マージン）を設定可能
- 水平・垂直をHキー、Vキーで切替
- ↑↓キー、shift + ↑↓での数値調整
- ガイドのプレビュー線を描画
- ガイド化時の色は設定不可（Illustratorの仕様）
- ガイドは「_guide」レイヤーに作成し、ロック

### note

https://note.com/dtp_tranist/n/n1085336d7265

### 更新履歴

- v1.0 (20250713) : 初期バージョン
- v1.1 (20250714) : レイヤー選択機能追加、リピート機能追加