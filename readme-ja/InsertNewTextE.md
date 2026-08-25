# InsertNewTextE


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--text--e.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-text-e.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewTextE.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

現在の表示領域の中心に、欧文フォントを適用したポイントテキストを1つ作成し、選択状態にします。作例やサンプルづくりのためのデモスクリプトです。

### 主な機能

- 表示領域の中心にポイントテキストを作成
- `TEXT_CONTENTS` の文言と `FONT_SIZE`（初期値 12pt）を適用
- `FONT_CANDIDATES` を上から順に試し、最初に見つかったフォントを適用
- 段落を中央揃えに設定
- 見た目の境界（visibleBounds）の中心を、表示領域の中心にそろえる
- 作成したテキストだけを選択状態にする

### 処理の流れ

1. ドキュメントが開いているか確認
2. 選択を解除
3. テキストフレームを作成して文言を設定
4. フォント候補・フォントサイズ・行揃えを適用
5. 見た目の中心を表示領域の中心にそろえる
6. 作成したテキストを選択

### メモ

- 文言・フォントサイズ・フォント候補は、スクリプト冒頭の「ユーザー設定」ブロックで変更できます。
- 位置合わせに `position` ではなく `visibleBounds` を使っているのは、`position` がベースライン基準のため、そのままでは見た目の中心とずれるためです。
- フォント候補がどれも見つからない場合は、アラートを出さずにドキュメントの現在のフォントのまま作成されます。
- 和文版は [InsertNewTextJ](InsertNewTextJ.md)、長めの欧文版は [InsertNewTextELong](InsertNewTextELong.md) を使ってください。

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### 更新履歴

- v1.1 (20260702) : 見た目の境界（visibleBounds）を基準に中心へ配置するよう変更
- v1.0 (20250401) : 初版
