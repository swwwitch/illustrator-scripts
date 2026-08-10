# TextCountStats

[![Direct](https://img.shields.io/badge/Direct%20Link-TextCountStats.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TextCountStats.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStats.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- Illustrator の選択テキストや全体の文字情報を統計的に可視化
- 文字数、段落数、行数、英単語数、全角文字数、半角カナ、種別（ポイント／エリア／パス上）、フォント数などをカウント
- 常駐パレット化：開いたまま選択を切り替え、［更新］で再集計

### 主な機能

- 選択したテキストオブジェクトの各種統計をパレット上に一覧表示
- 選択がない場合はドキュメント全体のテキストを対象に集計
- DOM 集計はメインエンジンへ BridgeTalk 委譲（常駐パレットの DOM 切断を回避）
- UI は日本語／英語対応

### 処理の流れ

1. 常駐パレットを表示（多重起動は自動で閉じてから再表示）
2. ［更新］押下（または表示直後）にメインエンジンへ集計を委譲
3. 戻り値（マーカー方式）を解析し、各パネルの値を更新

### 更新履歴

- v1.0 (20250806) : 初期バージョン
- v1.1 (20260702) : 常駐パレット化（#targetengine ＋ BridgeTalk 委譲）、更新ボタン・ステータス表示・ローカライズ整理

---

### スクリプト情報

- バージョン: v1.1
