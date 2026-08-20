# CenterAlignAsGroup


[![Direct](https://img.shields.io/badge/Direct%20Link-CenterAlignAsGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/CenterAlignAsGroup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterAlignAsGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

選択したオブジェクトを一時的にグループ化してから、整列パネルの［水平方向中央に整列］［垂直方向中央に整列］を実行します。選択内の位置関係を保ったまま、全体をまとめて中央へ動かせます。

整列パネルのコマンドはスクリプト（DOM）から直接呼び出せないため、アクション定義を一時ファイルに書き出して読み込み、実行後にすぐ破棄する「ダイナミックアクション」の方式を使っています。

### 主な機能

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
5. 複数選択ならグループ化 → アクションで中央揃え → グループ解除
6. ［字形の境界に整列］を元の状態へ戻す

### メモ

- 整列の基準（選択範囲／キーオブジェクト／アートボード）は［整列］パネルの設定に従います。**［選択範囲に整列］のままだと、グループ化によって対象が1つになるため何も動きません。** アートボードやキーオブジェクトを基準に設定してから実行してください。
- 複数のレイヤーにまたがる選択で中止するのは、グループ化すると全オブジェクトが最前面オブジェクトのレイヤーへ移動し、グループ解除しても元のレイヤーに戻らないためです。
- グループ化とグループ解除の分だけ、取り消し（command + Z）のステップが増えます。
- 連結（スレッド）されたテキストの文字を選択して実行すると、そのストーリーに属するすべてのテキストオブジェクトが対象になります。
- 天地中央だけ揃えたい（左右方向は動かしたくない）場合は [VerticalCenterAlignAsGroup](VerticalCenterAlignAsGroup.md) を使ってください。

### 更新履歴

- v1.0.0 (20260821) : 初版
