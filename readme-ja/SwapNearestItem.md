# SwapNearestItem

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapNearestItem.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SwapNearestItem.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItem.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトを基準に、指定方向（右／左／上／下）にある最も近いオブジェクトと自然な見た目で位置を入れ替えるIllustrator用スクリプトです。
- 複数選択時の特別処理や幅・高さ、隙間を考慮した入れ替えに対応します。

### 主な機能

- 上下左右方向での最短距離判定によるスワップ
- 幅・高さ、隙間を考慮した自然な位置調整
- 複数選択時の手動入れ替え処理対応
- 日本語／英語インターフェース対応

### 処理の流れ

1. ドキュメントと選択オブジェクトを確認
2. 指定方向に最も近いオブジェクトを検索
3. 高さ・幅・隙間を考慮して位置を入れ替え
4. 複数選択時は中心座標または端基準で入れ替え

### 更新履歴

- v1.0.0 (20250610) : 初版リリース
- v1.0.1 (20250612) : グループ・複合パスの除外、ロック対応追加
- v1.0.2 (20250613) : getCenter() と getSize() の導入による整理
- v1.0.3 (20250614) : 複数選択時の一時グループ処理追加
- v1.0.4 (20250615) : 一時グループ化処理の削除、整理

---
