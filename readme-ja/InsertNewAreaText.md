# InsertNewAreaText


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--areatext.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-areatext.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewAreaText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

現在の表示領域の中心に指定サイズの矩形を作り、それをエリア内文字に変換してサンプルテキストを流し込みます。フォント・フォントサイズ・行送り・行揃えを適用し、作成したエリア内文字だけを選択します。作例やサンプルづくりのためのデモスクリプトです。

### 主な機能

- 表示領域の中心に `AREA_WIDTH` × `AREA_HEIGHT`（初期値 220 × 110）の矩形を作成
- 矩形をエリア内文字に変換して `SAMPLE_TEXT` を流し込み
- `FONT_CANDIDATES`（秀英角ゴシック L → ヒラギノ角ゴ W3 → 源ノ角ゴシック）を上から順に試して適用
- フォントサイズ 12pt / 行送り 18pt / 水平・垂直比率 100% を適用
- 行揃えを均等配置（最終行左揃え）に設定
- プロポーショナルメトリクスを適用（`APPLY_PROPORTIONAL_METRICS`）
- 作成したエリア内文字だけを選択状態にする

### 処理の流れ

1. ドキュメントが開いているか確認
2. ロックされておらず表示されている最初のレイヤーを取得
3. 表示領域の中心に矩形を作成
4. 矩形をエリア内文字に変換してサンプルテキストを流し込み
5. 書式スタイルとフォントを適用
6. 作成したエリア内文字を選択

### メモ

- エリアのサイズ・サンプルテキスト・フォント候補・書式スタイルは、スクリプト冒頭の「ユーザー設定」ブロックで変更できます。
- 行揃えは `TEXT_STYLE.justification` に `Justification` のキー名（`FULLJUSTIFYLASTLINELEFT` / `CENTER` など）を文字列で指定します。存在しないキー名を書くとエラーになります。
- 編集可能なレイヤーが1つもない場合は、アラートを出さずに何もせず終了します。
- サンプルテキストの改行には `\r`（復帰）を使っています。ExtendScript から流し込む場合は `\n` ではなく `\r` が段落区切りになります。

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### 更新履歴

- v1.0 (20250813) : 初版
