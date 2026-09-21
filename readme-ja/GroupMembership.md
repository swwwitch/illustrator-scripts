# グループへの出し入れをダイアログで選んで実行

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupMembership.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/GroupMembership.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupMembership.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトを［既存のグループに入れる］［グループから出す］［1つずつグループにする］のどれで処理するかを、ダイアログで選んで実行します。
- 初期値は選択の状態から自動で決まります。グループの中のオブジェクトを選んでいれば［グループから出す］、グループ1つとほかのオブジェクトなら［既存のグループに入れる］、それ以外は［1つずつグループにする］です。
- ダイアログなしで1つの処理だけを実行する版もあります（[AddToGroup](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddToGroup.md) / [ReleaseFromGroup](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseFromGroup.md) / [GroupEachSelection](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEachSelection.md)）。

### 主な機能

- ［既存のグループに入れる］：一緒に選んだ既存のグループへ、ほかのオブジェクトを入れます。グループは解除しないので、不透明度・効果・名前・クリッピングマスクはそのまま残ります。重ね順も保たれます。グループが無いか複数あるときは、グループを1段解除してから1つのグループにまとめ直します（クリップグループは解除しません）。
- ［グループから出す］：グループの中で選んだオブジェクトを、グループの外へ出します。
- ［1つずつグループにする］：選択したオブジェクトを、1つずつ別々のグループにします。重ね順は変わりません。

### 使い方

1. 対象のオブジェクトを選択する（グループから出すときは、ダイレクト選択ツールなどでグループ内のオブジェクトを選ぶ）
2. スクリプトを実行し、処理を選んで［OK］をクリックする

### オプション

- ［出す位置］（［グループから出す］のとき）：［レイヤーの最前面］か［元のグループのすぐ前面］を選びます。［元のグループのすぐ前面］なら、ほかのオブジェクトとの重なりがほとんど変わりません。
- ［グループ化済みのものは飛ばす］（［1つずつグループにする］のとき）：選択の中のグループは、グループ化せずにそのまま残します。

### 注意点

- 選択が1つだけのときは［既存のグループに入れる］を、グループの中のオブジェクトを選んでいないときは［グループから出す］を選べません。
- 文字ツールで文字を選択しているときは実行できません。

### 紹介記事

- [【Illustrator】既存グループにオブジェクトを合流させたり、特定のオブジェクトを離脱させる｜DTP Transit 別館](https://note.com/dtp_tranist/n/n36fbd4162721)

### 更新履歴

- v1.0.0 (20260922) : 初版公開
