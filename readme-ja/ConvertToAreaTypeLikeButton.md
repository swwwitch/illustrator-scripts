# ConvertToAreaTypeLikeButton

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToAreaTypeLikeButton.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertToAreaTypeLikeButton.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButton.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行うツール。

選択内容に応じて自動でエリア内文字へ変換し、調整ダイアログを開く。

- ポイント文字 / パス上文字のみ → ボタン風（幅×1.2・高さ×1.6）に変換
- テキスト＋図形 → 図形をエリア内文字にしてテキストを流し込む
- エリア内文字のみ → そのまま調整ダイアログへ

### 調整ダイアログ

- フォントサイズの指定、「文字あふれ解消」で枠に収まる最大サイズへ自動縮小
- フレームサイズ（幅・高さ）の変更
- 左右インデント（連動可）と外側からの間隔
- 行揃え・テキストの配置は常に中央
- プレビューは常時ONで、変更を即座に反映

### 補足

- 縦方向の中央配置にはフレーム整列アクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

### スクリプト情報

- バージョン: v1.0.0
