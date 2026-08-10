# PathCleanupTool-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-PathCleanupTool-v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/PathCleanupTool-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathCleanupTool-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2026-03-20
- 選択したパス（グループ／複合パス内も含む）の冗長アンカー・同座標アンカー・直線扱い可能なベジェハンドルを削除して最適化
- 「その他」タブからスムーズ化／コーナー化／アンカー追加／アンカー分割も実行可能

### 主な機能

- 直線上で冗長なアンカーポイントを削除
- 同じ座標のアンカーポイントを削除
- 直線として扱えるベジェ区間のハンドルを削除
- ロック／非表示オブジェクト（親・レイヤー含む）は自動スキップ
- ダイアログ表示時点の選択を固定し、情報表示と実行対象を一致
- 許容誤差はアンカー削除用とハンドル削除用で個別に調整可能
- スムーズ化は前後アンカーが同一点や極端に近い場合のガード付き
- オープンパス端点は循環参照せず自然な接線方向を使用
- 角度差・長さ差が大きい場合はハンドル長を抑えて破綻を抑制
- 実処理中の例外は UI 系の保存復元と分離して最小限ログ出力

### 更新履歴

- v1.4.1 (2026-03-20) : 現行版

### スクリプト情報

- バージョン: v1.4.1
