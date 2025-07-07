# 選択したグループ内のサブグループを解除して、グループ構造を簡素化

[![Direct](https://img.shields.io/badge/Direct%20Link-SimplifyGroups.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SimplifyGroups.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したグループ内のサブグループを再帰的に解除します。
- 最外層のグループは解除せず残します。
- グループと非グループが混在している場合はまとめて1つのグループにします。

![](https://www.dtp-transit.jp/images/ss-1464-1026-72-20250707-162428.png)

### 主な機能

- サブグループの再帰的解除
- 混在選択時の自動グループ化
- Illustratorメニュー「グループ解除」コマンドの利用

### 処理の流れ

1. ドキュメントと選択を確認
2. グループと非グループが混在していればまとめてグループ化
3. グループ内のサブグループを再帰的に探索し解除

### 更新履歴

- v1.0.0 (20250707) : 初版公開