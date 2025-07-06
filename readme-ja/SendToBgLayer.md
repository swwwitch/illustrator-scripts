# 選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック

[![Direct](https://img.shields.io/badge/Direct%20Link-SendToBgLayer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SendToBgLayer.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要：

選択されたオブジェクトを「bg」レイヤーに重ね順を維持したまま移動し、最背面に配置します。

![](https://assets.st-note.com/img/1675146941932-Nxauo7hIWW.png?width=1200)

### 処理の流れ：
  1. ドキュメントと現在のアクティブレイヤーを取得
  2. 「bg」レイヤーの存在を確認し、なければ作成
  3. 選択されたオブジェクトを重ね順を維持したまま「bg」レイヤーに移動（エラーを無視）
  4. 「bg」レイヤーを最背面に移動してロック
  5. 元のレイヤーを再アクティブ化

### 対象：

選択されたオブジェクト

### 除外：

未選択時には何もしない

### 更新日：

2025-06-24

### ブログ

- [Illustratorで「選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック」までを1ストロークで実行する｜DTP Transit 別館](https://note.com/dtp_tranist/n/nf7c1e8a0f0c7)