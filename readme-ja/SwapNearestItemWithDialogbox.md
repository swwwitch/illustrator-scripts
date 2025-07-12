# オブジェクトを入れ替え

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapNearestItem.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SwapNearestItem.jsx)


[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)


---

- 選択中のオブジェクトと、指定方向（上下左右）にある最も近いオブジェクトの位置を入れ替える
- ダイアログボックスを表示中、↑↓キーで上下、←→キーで左右のオブジェクトを入れ替える

### 概要：
選択された1つのオブジェクトを起点に、指定方向（右／左／上／下）にある最も近いオブジェクトと
見た目上の自然な位置で入れ替えるスクリプトです。

### 処理の流れ：
1. ドキュメントと選択状態をチェック
2. 指定方向にある最も近いオブジェクトを検索
3. 高さや幅、隙間を考慮して自然な位置に2つのオブジェクトを入れ替え

### 対象オブジェクト：
PathItem, CompoundPathItem, GroupItem, TextFrame, PlacedItem, RasterItem,
SymbolItem, MeshItem, PluginItem, GraphItem

### 除外オブジェクト：
LegacyTextItem, NonNativeItem, Guide, Annotation, PathPoint、
ロック／非表示のオブジェクトやレイヤー、グループや複合パスの内部構造

### 初版作成日

2025-06-10

更新履歴：
- 1.0.0 初版リリース
- 1.0.1 グループや複合パスの子要素を除外、レイヤーロック対応
- 1.0.2 getCenter() と getSize() の導入によりロジックを整理
- 1.0.3 複数選択時の一時グループ化と解除に対応
- 1.0.4 複数選択時の一時グループ化・解除処理を削除
