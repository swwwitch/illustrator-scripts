# 画像をパスごとに切り分け

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipImageWithEachPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipImageWithEachPath.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipImageWithEachPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。

### 主な機能

- パス1本につき画像を1つ複製し、そのパスでクリッピングマスクを作ります。
- クリップグループは、マスクに使ったパスがあった重ね順に入ります。
- マスクには長方形・楕円などのパスと複合パスが使えます。
- 配置画像（リンク）と埋め込み画像のどちらにも使えます。
- 元の画像は最後に削除し、作ったクリップグループを選択した状態で終わります。

### 使い方

1. 画像の上に、切り出したい範囲を囲むパスを描きます（いくつでも）。
2. 画像1つとパスをまとめて選択します。
3. スクリプトを実行します。ダイアログボックスは出ません。

### 注意点

- 選択に含められる画像は1つだけです。
- 画像と重ならないパスは、マスクにすると中身が空になるので、何もせずそのまま残します。
- パスの塗りと線は、クリッピングマスクにしたときに消えます。
- 選択の中のテキスト・グループ・ガイドは対象外です。グループの中にあるパスも使いません。

### 更新履歴

- v1.0.0 (2026-11-03) : 初版
