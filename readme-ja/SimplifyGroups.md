# 選択したグループ内のサブグループを解除して、グループ構造を簡素化

[![Direct](https://img.shields.io/badge/Direct%20Link-SimplifyGroups.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SimplifyGroups.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SimplifyGroups.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したグループ内のサブグループを再帰的に解除します。
- 最外層のグループは解除せず残します。
- グループ以外のオブジェクトも選んでいるときは、それらをグループに取り込みます。グループが複数あるときは、全体をまとめて1つのグループにします。

![](https://www.dtp-transit.jp/images/ss-1464-1026-72-20250707-162428.png)

### 主な機能

- サブグループの再帰的解除
- グループ以外のオブジェクトの取り込み（グループが複数なら自動グループ化）
- Illustratorメニュー「グループ解除」コマンドの利用

### 処理の流れ

1. ドキュメントと選択を確認
2. グループ以外のオブジェクトがあれば、グループが1つならそこへ取り込み、複数ならまとめてグループ化
3. グループ内のサブグループを再帰的に探索し解除

### 注意点

- ロックまたは非表示のサブグループは解除せずに残します。

### 更新履歴

- v1.0.0 (20250707) : 初版公開
- v1.3.1 (20260922) : 処理を整理。テキストでマスクしたクリップグループへの取り込みに対応。ロック・非表示のサブグループは解除せずに残すように変更