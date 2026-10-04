# 基準サイズと倍率からタイプスケールを生成

[![Direct](https://img.shields.io/badge/Direct%20Link-TypeScaler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TypeScaler.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeScaler.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 基準フォントサイズと倍率からタイプスケールを自動生成
- 選択テキストにフォントサイズを適用、または見本を生成
- プレビュー機能でサイズを即時確認可能

### 主な機能

- 倍率に基づいたサイズリスト生成
- 選択テキストへのフォントサイズ適用
- 複数サイズの見本テキスト生成
- 単位に応じたラベル表示

### 処理の流れ

1. ダイアログで基準サイズと倍率を入力
2. 自動生成されたサイズリストを表示
3. 選択したサイズを適用または見本を作成

### 参考

https://note.com/hiro_design_n/n/nc95a1d2d86a4

### 更新履歴

- v1.0 (20250728) : 初期バージョン
- v1.1 (20250729) : UI改善とローカライズ対応
- v1.2 (20250801) : ダイアログボックスを再度開いたときに値を記憶する機能を追加
- v1.3.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.3.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.2 (20260929) : ダイアログの不透明度を98%に変更
- v1.3.3 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.3.4（2026-10-01）ボタン行を共通部品にそろえ、ダイアログの位置は共通部品に任せるようにした。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.3.5（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.3.6（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

### スクリプト情報

- バージョン: v1.3.6
