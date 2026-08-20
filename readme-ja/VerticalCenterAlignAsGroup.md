# VerticalCenterAlignAsGroup


[![Direct](https://img.shields.io/badge/Direct%20Link-VerticalCenterAlignAsGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/VerticalCenterAlignAsGroup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VerticalCenterAlignAsGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

選択したオブジェクトを一時的にグループ化してから、整列パネルの［垂直方向中央に整列］（天地中央）を実行します。左右方向は動かさないため、横位置を保ったまま天地だけを揃えられます。

整列パネルのコマンドはスクリプト（DOM）から直接呼び出せないため、アクション定義を一時ファイルに書き出して読み込み、実行後にすぐ破棄する「ダイナミックアクション」の方式を使っています。

### 主な機能

- 天地中央のみ整列（左右方向の揃えは実行しない）
- 選択が2つ以上のときは一時的にグループ化して整列し、終了後にグループ解除
- 実行中だけ［字形の境界に整列］（ポイント文字・エリア内文字）をONにし、終了時に元の状態へ復元
- 文字ツールで文字を選択している場合は、そのテキストオブジェクトを対象に切り替え
- 複数のレイヤーにまたがる選択は、アラートを表示して中止
- 読み込んだアクションは実行後に必ず削除（［アクション］パネルに残りません）

### 処理の流れ

1. ドキュメントと選択状態を確認
2. 文字を選択している場合は、テキストオブジェクトを選択し直す
3. 選択が複数レイヤーにまたがっていないか確認
4. ［字形の境界に整列］の現在値を退避してONにする
5. 複数選択ならグループ化 → アクションで天地中央揃え → グループ解除
6. ［字形の境界に整列］を元の状態へ戻す

### メモ

- 整列の基準（選択範囲／キーオブジェクト／アートボード）は［整列］パネルの設定に従います。**［選択範囲に整列］のままだと、グループ化によって対象が1つになるため何も動きません。** アートボードやキーオブジェクトを基準に設定してから実行してください。
- 複数のレイヤーにまたがる選択で中止するのは、グループ化すると全オブジェクトが最前面オブジェクトのレイヤーへ移動し、グループ解除しても元のレイヤーに戻らないためです。
- グループ化とグループ解除の分だけ、取り消し（command + Z）のステップが増えます。
- 連結（スレッド）されたテキストの文字を選択して実行すると、そのストーリーに属するすべてのテキストオブジェクトが対象になります。
- 左右方向もまとめて中央に揃えたい場合は [CenterAlignAsGroup](CenterAlignAsGroup.md) を使ってください。

### 更新履歴

- v1.0.0 (20260821) : 初版
