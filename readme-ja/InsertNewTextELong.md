# InsertNewTextELong


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--text--e--long.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-text-e-long.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewTextELong.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

現在の表示領域の中心に、欧文フォントを適用した長めのポイントテキストを1つ作成し、中央揃えにして選択状態にします。作例やサンプルづくりのためのデモスクリプトです。

### 主な機能

- 表示領域の中心にポイントテキストを作成
- `TEXT_CONTENTS` の文言（初期値 `Design with clarity, build with intent.`）と `FONT_SIZE`（初期値 12pt）を適用
- `FONT_CANDIDATES` を上から順に試し、最初に見つかったフォントを適用
- 段落を中央揃えに設定
- 作成したテキストだけを選択状態にする

### 処理の流れ

1. ドキュメントが開いているか確認
2. 選択を解除
3. 表示領域の中心座標を取得
4. テキストフレームを作成して文言を設定
5. 先にフォントを確定してから、フォントサイズ・行揃えを適用
6. 中心へ配置し、作成したテキストを選択

### メモ

- [InsertNewTextE](InsertNewTextE.md) との違いは文言の長さと配置方法です。こちらは `position` と `width` / `height` で中心へ移動します。
- 文言・フォントサイズ・フォント候補は、スクリプト冒頭の「ユーザー設定」ブロックで変更できます。
- フォント候補がどれも見つからないときはアラートを表示し、そのまま現在のフォントで作成を続けます。
- 和文版は [InsertNewTextJ](InsertNewTextJ.md) を使ってください。

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### 更新履歴

- v1.1 (20250813) : 設定値とフォント候補の外出し、ユーティリティ関数への分割、段落中央揃えを追加
- v1.0 (20250401) : 初版
