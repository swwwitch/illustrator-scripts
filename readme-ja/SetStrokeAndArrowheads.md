# SetStrokeAndArrowheads

[![Direct](https://img.shields.io/badge/Direct%20Link-SetStrokeAndArrowheads.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SetStrokeAndArrowheads.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SetStrokeAndArrowheads.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトの線幅と矢印（始点／終点の形状・倍率・先端位置）をまとめて設定します。
- 矢印は Illustrator の DOM から操作できないため、一時アクション（ai_plugin_setStroke）を生成して実行します。
- ダイアログでプレビューできます。線幅は DOM で即時反映、矢印を含む設定はアクション実行＋取り消しで反映します。

### 処理の流れ

1. ドキュメントと選択オブジェクトの有無を確認
2. ダイアログで線幅・矢印・先端位置を入力（プレビュー可）
3. 入力値から .aia（アクション）ソースを生成
4. 一時ファイルとして書き出し → 読み込み → 実行 → 破棄

### 注意点

- 矢印名・先端位置名は Illustrator の UI 表示名と一致している必要があります（言語に依存）。
- 矢印の倍率キー（asc1 / asc2）は推定値です。

---

### スクリプト情報

- バージョン: v1.0.0
- 初回リリース: 2026-07-22
- 最終更新: 2026-07-22
