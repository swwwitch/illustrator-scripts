# ConvertToRectangle

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/ConvertToRectangle.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2026-05-24
- 選択オブジェクトの境界に合わせて長方形を作成する Illustrator スクリプト
- 作成単位は「オブジェクトごと」または「選択範囲全体」から選択
- マージン（定規単位、負値で内側）、角丸ライブエフェクト、塗り・線プリセットを指定して生成
- 元オブジェクトは「残す」「クリッピングマスクにする」「削除」から選択可能
- プレビュー対応。ダイアログ表示中は選択を一時的に 50% にディムして比較しやすく

### 主な機能

- プレビュー境界での計測、テキストのアウトライン化計測（テキスト選択時のみ有効）、定規単位に連動したマージン（負値で内側）を指定可能
- ダイアログ表示中は選択オブジェクトの不透明度を一時的に 50% へ下げてプレビューしやすく（不透明度を変えるプリセット選択時は自動で無効化）
- 塗り・線プリセット（線幅は Illustrator の「線幅」単位に連動）、角丸ライブエフェクト（定規単位の半径）、重ね順、元オブジェクトの扱い（残す／クリッピングマスクにする／削除）を設定可能
- 「クリッピングマスクにする」はリンク画像／埋め込み画像が選択されているときのみ選択可
- クリッピングマスク化では元オブジェクトの重ね順を維持
- 「選択範囲全体」では前面／背面に応じて最前面／最背面の項目を基準
- ロック・非表示オブジェクトは対象外
- 実行中はバウンディングボックスとエッジ表示を一時切替し、終了時（OK／キャンセル／エラー）に元へ復元

### キーボード操作

- マージン・角丸半径入力：↑↓ で ±1、Shift+↑↓ で ±10、Alt+↑↓ で ±0.1
- ショートカット：S/G（作成単位）、P/O/A/T（オプション）、F/B（重ね順）、N/M/D（元オブジェクト）

### 更新履歴

- v1.0.1 (2026-05-20) : 初版
- v1.1.0 (2026-05-24) : 単位換算情報を集約し、表示切替の復元処理を安全化

### スクリプト情報

- バージョン: v1.1.0
