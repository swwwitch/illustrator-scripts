# RectangleToArc

[![Direct](https://img.shields.io/badge/Direct%20Link-RectangleToArc.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/RectangleToArc.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArc.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2026-05-19
- 選択した長方形を、左下隅・上辺中央・右下隅の 3 点を通る円弧に変換
- 長方形の幅と高さから円の半径・中心・開始角・終了角を求め、90 度以内のベジェセグメントに分割して円弧を生成
- 元の長方形は削除

### 主な機能

- 閉じた 4 点パスかつ各辺が水平／垂直の長方形のみ処理（回転長方形・台形・菱形・不定形は対象外）
- 生成する円弧は塗りなし・線あり
- 元長方形に線がある場合は線の色・太さを引き継ぎ
- 元長方形が線なしでも変換は実行
- 複数選択に対応

### 更新履歴

- v1.0.1 (2026-05-19) : 現行版

### スクリプト情報

- バージョン: v1.0.1
