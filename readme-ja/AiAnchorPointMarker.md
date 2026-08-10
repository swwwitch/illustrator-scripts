# AiAnchorPointMarker

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAnchorPointMarker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/AiAnchorPointMarker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAnchorPointMarker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトの全アンカーポイントに、マーカー（正方形／最前面オブジェクト／シンボル）を配置するユーティリティです。ダイアログを閉じずにライブプレビューしながら設定できます。

- 追加するオブジェクトを選択：アンカーポイントを自動生成（正方形）／最前面のオブジェクトを複製／ドキュメント内シンボルのインスタンス
- 正方形は「大きさ（pt・小数可、⌘＋↑↓で±0.1）」「カラー（自前のRGBダイアログ）」「シンボル化（既定ON・シンボル名『アンカーポイント』）」を指定
- オプション：スケール（%、最前面オブジェクト・シンボルに適用）／レイヤーに移動（_anchorpoint）／グループ化（既定ON）／9軸の基準点（マーカーをアンカーに合わせる位置）
- 自動生成時はスケール・9軸をディム表示し、基準点は中央に固定
- ライブプレビューは専用レイヤーに描画し、OK／キャンセルで確実に片付け
- 実行中はエッジ表示とライブコーナー注釈を一時的に隠す（開始・終了でトグル）
- 日英ローカライズ、ライト／ダークUIに追従

### 更新履歴

- v1.0.0: 初期バージョン。マーカー配置（正方形／最前面／シンボル）、基準点（9軸）、スケール、レイヤー移動、グループ化、ライブプレビュー。
- v1.0.1: 既定のアンカーポイントカラーを RGB(79,128,255) に変更。

### スクリプト情報

- バージョン: v1.0.1
