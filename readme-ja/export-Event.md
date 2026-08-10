# export-Event

[![Direct](https://img.shields.io/badge/Direct%20Link-export-Event.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/export-Event.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export-Event.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

  アクティブドキュメントの全アートボードを、名前ごとのルールで PNG 書き出しします。

### 書き出しルール

  - "title" / "title2"（"-..." 付きを含む） … 背景「透明」で 100% 書き出し
  - "Doorkeeper" … 背景「白」で 100% と 200% の 2 種類書き出し（200% は "-200" を付加）
  - "シンボル一覧" … 書き出し対象外
  - 上記以外 … 背景「白」で 100% 書き出し

  ルールを追加・変更する場合は `buildExportJobs()` を編集してください。
  空配列を返せば「除外」、複数要素を返せば「複数倍率の書き出し」になります。

### 出力

  - 保存先 … ドキュメントと同じフォルダ
  - ファイル名 … "<ドキュメント名>-<アートボード名>[suffix].png"
  - 書き出し後は macOS のみ Finder で保存先を自動オープン

### 更新履歴

  - 2025-04-22 … 初版
  - 2026-06-03 … 書き出し前に "Guides Preview for Trim View" レイヤー（"*" 付き含む）を非表示にし、書き出し後に再表示 / 未保存ドキュメントのガードを追加
  - 2026-06-15 … "Guides Preview for Trim View" レイヤー（"*" 付き含む）を非表示ではなく書き出し前に削除するよう変更
  - 2026-06-17 … 定数の巻き上げで削除が機能していなかった不具合を修正。仕様を再び「書き出し前に非表示 → 書き出し後に再表示」に戻す（中断・エラー時も再表示）
  - 2026-07-03 … 進捗ウィンドウを native progressbar に戻す / "Guides Preview for Trim View" レイヤーの非表示・再表示処理を削除

### スクリプト情報

- バージョン: v1.0.4
