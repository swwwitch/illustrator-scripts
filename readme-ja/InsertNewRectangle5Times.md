# InsertNewRectangle5Times


[![Direct](https://img.shields.io/badge/Direct%20Link-insert--new--rectangle--5times.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/demo/insert-new-rectangle-5times.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewRectangle5Times.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

現在の表示領域の中心付近に、黒く塗ってランダムな不透明度を与えた正方形を5つ作成し、互いに重ならないように配置してから［シェイプに変換］（Convert to Shape）と［ピクセルグリッドに最適化］（Make Pixel Perfect）を適用します。作例やサンプルづくりのためのデモスクリプトです。

### 主な機能

- `RECT_SIZE`（初期値 100）四方の正方形を `RECT_COUNT`（初期値 5）個作成
- ドキュメントのカラーモードを判定し、CMYK なら K100、RGB なら R0/G0/B0 で塗る
- 不透明度を `OPACITY_MIN`〜`OPACITY_MAX`（初期値 30〜100%）の範囲でランダムに設定
- 表示領域の中心に同サイズ・同位置の長方形が残っていれば削除
- ランダムに位置を振り直し、`padding`（初期値 6）以上の間隔を空けて重ならないように配置
- 収まらない場合は探索範囲を段階的に広げて再試行（最大 `maxScaleFactor` 倍）
- 作成した5つを選択し、［シェイプに変換］［ピクセルグリッドに最適化］を実行

### 処理の流れ

1. ドキュメントが開いているか確認
2. 作成先レイヤーを決定（必要ならアクティブレイヤーを切り替え）
3. 表示領域の中心に残っている既存の長方形を削除
4. 中心付近にばらつかせて5つ作成し、塗り・線・不透明度を設定
5. 重ならない位置が見つかるまでランダムに配置し直す
6. 作成した5つを選択して［シェイプに変換］［ピクセルグリッドに最適化］を実行

### メモ

- 個数・サイズ・不透明度の範囲・重なり回避の探索条件は、スクリプト冒頭の「ユーザー設定」ブロックで変更できます。
- 手順3では、表示領域の中心にある「同じサイズ・同じ位置」の長方形を削除します。[InsertNewRectangle](InsertNewRectangle.md) を実行した直後に続けて使うことを想定した処理ですが、条件に合致する既存オブジェクトがあると一緒に削除されます。
- `attemptsPerItem` 回の試行で置けなかった場合は探索範囲を1段広げてやり直します。`maxScaleFactor` 倍まで広げても置けなかった場合は、最後に試した位置のままになります。
- ランダム値を使うため、実行するたびに配置と不透明度が変わります。
- 1つだけ作成する版は [InsertNewRectangle](InsertNewRectangle.md) を使ってください。

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/n509eb6aa0a19)

### 更新履歴

- v1.3 (20260305) : 重なりを回避して配置するように変更
- v1.2 (20250713) : 関数整理とヘッダー更新
- v1.1 (20250511) : コメント整理とロジック改善
- v1.0 (20250401) : 初版
