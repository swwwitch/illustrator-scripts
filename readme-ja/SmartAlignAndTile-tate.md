# SmartAlignAndTile-tate

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAlignAndTile--tate.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignAndTile-tate.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile-tate.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2026-02-26
- 選択したオブジェクトを縦方向に整列し、指定した間隔と横方向の数（列数）で再配置するスクリプト。
- プレビュー時に境界線を含むオプション、ランダム配置、単位の自動取得、上下キーでの数値変更に対応。
- プレビュー時にUndo履歴を汚さないように管理し、OK時は1回のUndoで取り消せるように確定します。

### 主な機能

- 縦方向整列と再配置
- 列数（横方向の数）指定
- ランダム配置オプション
- プレビュー時の境界含む切替
- 単位自動対応
- キーボードで間隔・列数調整
- Undoを汚さないプレビューと一括取り消し（1回のUndo）

### 処理の流れ

- オブジェクト選択確認
- ダイアログ表示（各オプション設定）
- プレビュー更新
- 実行時に配置確定

### オリジナルアイデア

John Wundes
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

### 更新履歴

- v1.8 (20260226) : 縦方向前提（列数指定）にロジックとUIを変更、変数/関数名を整理
- v1.7 (20260119) : プレビュー時にUndo履歴を汚さないように管理し、OK時は1回のUndoで取り消せるように確定
- v1.6 (20250809) : 「プレビュー境界を使用」をOFFのとき geometricBounds を使用するように調整
- v1.0 (20250716) : 初期バージョン
- v1.1 (20250717) : 安定性改善、行数ロジック修正
- v1.2 (20250718) : コメント整理、ローカライズ統一、ランダム基準位置補正改善
- v1.3 (20250801) : グリッド機能の追加、ガターを縦横個別に設定
- v1.4 (20250801) : ローカライズを調整
- v1.5 (20250802) : 横／縦の連動機能を追加、プレビュー境界を使用のロジックを調整
- v1.6 (20250809) : 「プレビュー境界を使用」をOFFのとき geometricBounds を使用するように調整

---

### スクリプト情報

- バージョン: v1.8
