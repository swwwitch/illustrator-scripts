# 線幅と矢印をまとめて設定

[![Direct](https://img.shields.io/badge/Direct%20Link-SetStrokeAndArrowheads.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/single-function/SetStrokeAndArrowheads.jsx)

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

### 更新履歴

- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (20260928) : 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした。失敗時の警告を「アクションを実行できませんでした。」に変更
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした

### スクリプト情報

- バージョン: v1.1.2
- 初回リリース: 2026-07-22
- 最終更新: 2026-09-27
