# 既存のグループにオブジェクトを加える

[![Direct](https://img.shields.io/badge/Direct%20Link-AddToGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/single-function/AddToGroup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddToGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトを、一緒に選んだ既存のグループへ加えます。
- グループは解除しないので、不透明度・効果・名前・クリッピングマスクはそのまま残ります。クリップグループに加えたオブジェクトはマスクされます。
- 重ね順は保たれます。グループより前面にあったものはグループ内の最前面へ、背面にあったものは最背面へ入ります。
- グループが無いか複数あるときは、グループを1段解除してから1つのグループにまとめ直します。クリップグループは解除せずにそのまま入れます。

### 使い方

1. まとめたいオブジェクトと、追加先のグループを一緒に選択する
2. スクリプトを実行する

### 注意点

- 選択が2つ未満のとき、文字ツールで文字を選択しているときは何もしません。
- ダイアログはありません。

### 紹介記事

- [【Illustrator】既存グループにオブジェクトを合流させたり、特定のオブジェクトを離脱させる｜DTP Transit 別館](https://note.com/dtp_tranist/n/n36fbd4162721)

### 更新履歴

- v1.0.0 (20260306) : 初版公開
- v1.0.2 (20260922) : グループを解除せずにオブジェクトを加えるように変更（クリップグループのマスクや、グループの効果・名前が残る）。文字ツールで文字を選択しているときは何もしないように修正
