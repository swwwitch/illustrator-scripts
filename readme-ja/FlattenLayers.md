# FlattenLayers

[![Direct](https://img.shields.io/badge/Direct%20Link-FlattenLayers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenLayers.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 作成日

2025-04-14

### 更新日

2026-04-15

### 概要

- 実行前にダイアログを表示し、処理条件を選択可能
- 除外レイヤーを残し、それ以外のレイヤー／サブレイヤー配下のオブジェクトを指定レイヤーへ移動してフラット化
- ロック / 非表示のレイヤー・オブジェクトをそれぞれ対象外にするか選択可能
- ロックレイヤー / 非表示レイヤー / 非表示オブジェクト / ガイド / サブレイヤーが存在しない場合、対応するUIを自動的にディム表示
- 対象外設定の一括切替は、有効な項目だけを一括で ON / OFF
- 必要に応じて、中身が残ったサブレイヤーを上位レベルのレイヤーへ移動可能
- ガイドの扱いを「統合」「現在のレイヤーに保持」「別レイヤーに移動」から選択可能
- 「現在のレイヤーに保持」では、サブレイヤー直下のガイドを1つ上のレイヤーへ繰り上げて保持
- 「別レイヤーに移動」では、統合後のガイドを指定レイヤーへ移動
- 空のレイヤー／サブレイヤーを、削除可能なものがなくなるまで再帰反復で削除
- skipLockedLayers / skipHiddenLayers はトップレベルレイヤーとサブレイヤーの両方に一貫適用
- ガイドだけ残っているレイヤーは空レイヤーとは見なさない
- まとめ先のレイヤー名、既存まとめ先の再利用、レイヤーカラーを指定可能
- 既存まとめ先を再利用しない場合、同名レイヤーが存在すれば連番付きの別名で新規作成

### 処理の流れ

1. ドキュメント取得（未オープンなら終了）
2. ダイアログで処理条件を選択（該当しない項目は自動的にディム表示、キャンセル時は終了）
3. 設定に応じて既存まとめ先レイヤーを取得、または一意な名前で新規作成
4. 除外レイヤーを除いて、条件に合う全レイヤー／サブレイヤー配下のオブジェクトをまとめ先レイヤーへ移動してフラット化
5. ガイドモードが「現在のレイヤーに保持」の場合、サブレイヤー直下のガイドを上位レイヤーへ繰り上げ
6. 必要に応じて、中身が残ったサブレイヤーを上位レベルのレイヤーへ移動
7. ガイドモードが「別レイヤーに移動」の場合、統合後のガイドを指定レイヤーへ移動
8. 設定が ON の場合、空のレイヤー／サブレイヤーがなくなるまで再帰反復で削除
9. 失敗件数があれば、処理後に件数だけを簡潔に通知

### 更新履歴

- v1.0 (20250414) : 初期バージョン
- v1.7.1 (20260408) : 対象外UIの自動ディム表示、一括切替の有効項目限定、skipLockedLayers / skipHiddenLayers のサブレイヤーまでの一貫適用、ガイド panel 初期化の明確化、概要とコメントの更新
- v1.7.3 (20260413) : 除外レイヤー名を「bg」のみに変更（「背景」「background」を削除）
- v1.7.4 (20260415) : 除外レイヤー内のガイドをガイドレイヤーへ移動するオプションを追加（「別レイヤーに移動」選択時、デフォルトON）

Illustrator script to flatten layers. It keeps excluded layers (bg),
moves objects under all other layers and sublayers into a specified destination layer, optionally
promotes remaining non-empty sublayers to the top level, and recursively deletes layers that become empty.
Locked / hidden layers and objects can be excluded independently, and dialog items with no matching targets
are dimmed automatically. Guides can be handled in one of three modes: integrate, keep in the current layer,
or move to another layer.

### スクリプト情報

- バージョン: v1.7.4
