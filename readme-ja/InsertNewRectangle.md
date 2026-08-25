# InsertNewRectangle


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--rectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-rectangle.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

現在の表示領域の中心に黒く塗った正方形を1つ作成して選択し、［シェイプに変換］（Convert to Shape）と［ピクセルグリッドに最適化］（Make Pixel Perfect）を適用します。作例やサンプルづくりのためのデモスクリプトです。

### 主な機能

- 表示領域の中心に `RECT_SIZE`（初期値 100）四方の正方形を作成
- ドキュメントのカラーモードを判定し、CMYK なら K100、RGB なら R0/G0/B0 で塗る
- 線（ストローク）はなしに設定
- アクティブレイヤーがロック中／非表示のときは、ロック解除かつ表示されている最初のレイヤーに切り替え
- 作成した正方形だけを選択し、［シェイプに変換］［ピクセルグリッドに最適化］を実行

### 処理の流れ

1. ドキュメントが開いているか確認
2. 作成先レイヤーを決定（必要ならアクティブレイヤーを切り替え）
3. 表示領域の中心座標を取得
4. 正方形を作成して塗りと線を設定
5. 作成した正方形だけを選択
6. ［シェイプに変換］［ピクセルグリッドに最適化］を実行

### メモ

- 正方形の大きさは、スクリプト冒頭の「ユーザー設定」ブロックの `RECT_SIZE` で変更できます。
- ロック解除かつ表示されているレイヤーが1つもない場合は、アラートを表示して中止します。
- ［ピクセルグリッドに最適化］は座標を整数に丸めるため、表示倍率によっては中心から最大 0.5pt ほどずれます。
- 5つまとめて作成する版は [InsertNewRectangle5Times](InsertNewRectangle5Times.md) を使ってください。

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### 更新履歴

- v1.2 (20250713) : 関数への分割とヘッダー整理
- v1.1 (20250511) : コメント整理とロジック改善
- v1.0 (20250401) : 初版
