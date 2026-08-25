# AiQuickPrefsPalette-simple

[![Direct](https://img.shields.io/badge/Direct%20Link-AiQuickPrefsPalette--simple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/AiQuickPrefsPalette-simple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiQuickPrefsPalette-simple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

Illustrator の各種環境設定の切り替えと、選択オブジェクトの反転・回転を、常駐パレットでまとめて操作するユーティリティです。操作した時点で即時反映されます。

- パレット（常駐エンジン）で表示し、書き込み・DOM 操作は BridgeTalk でメインエンジンへ委譲（読み出しは同期で直接取得）
- 上段は2カラム（左:キー増加／整列オプション／変形オプション・右:反転と回転／字形の境界に整列）、その下に全幅でコピー/ペースト・描画
- 反転・回転はアイコンボタンで実行し、9軸（3×3）ウィジェットで基点を指定。アイコンはライト／ダーク UI に合わせて配色を自動切り替え
- 環境設定ダイアログ等の外部変更は、パレットをクリック（再アクティブ）で同期
- パレットがアクティブなとき esc キーで閉じる

### パネルと項目

- キー増加：カーソル移動量（cursorKeyLength）。単位ポップアップで定規単位を切替、↑↓ / Shift / Option で増減
- 整列オプション：プレビュー境界
- 字形の境界に整列：ポイント文字／エリア内文字
- 変形オプション：パターン／角／線幅と効果
- 変形：左右反転／上下反転／回転（反時計回り・時計回り）をアイコンボタンで実行。9軸（3×3）の基準点ウィジェットで反転・回転の基点を指定（既定は中央）。基点は選択全体の可視バウンディングを基準に算出。アイコンはライト／ダーク UI に合わせて配色を自動切り替え
- コピー/ペースト：書式なしペースト／コピー元のレイヤーにペースト
- 描画：リアルタイムの描画と編集／プレビュー更新（GPU プレビューを更新）

### 紹介記事（note）

https://note.com/dtp_tranist/n/n41d8dc1961be

### 更新履歴

- v2.0.0 (20260630): 「反転と回転」パネルを FlipRotatePalette のアイコン UI へ差し替え（左右反転／上下反転／回転 CCW・CW のアイコンボタン＋9軸の基準点ウィジェット、ライト／ダーク対応）。反転・回転の基点を選択中心固定から 9 軸の任意基準点に変更（btTransformSelection＋getAnchorExpressions に共通化、適用後に app.redraw()）。方向ラジオと水平／垂直／45°回転のテキストボタンを廃止。
- v1.0 (20250804): 初期バージョン。

### スクリプト情報

- バージョン: v2.0.0
