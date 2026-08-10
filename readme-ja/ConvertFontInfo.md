# ConvertFontInfo

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertFontInfo.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ConvertFontInfo.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertFontInfo.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したテキストをフォント情報に変換するIllustrator用スクリプトです。
- ダイアログで変換形式を選び、対象テキストを即時プレビューしながら書き換えます。
- 右列には、選択中テキストの実フォント情報をもとにした変換結果を表示します。
- OK時は選択中の形式を再適用し、キャンセル時は開始時の内容、フォント、サイズ、行揃えに戻します。

### 主な機能

- フォント名、スタイル、フォント名＋スタイル、PostScript名、フルネーム＋サイズ、詳細表示を選択可能
- 初期状態は「フォント名＋スタイル」を選択
- F/S/B/P/M/Dキーで変換形式を切り替え
- Enter／Returnキーで確定（OKがデフォルトボタン）
- 詳細表示ではラベル行と値行を分け、行揃えを左揃えに変更
- 選択したテキストフレームを保持し、プレビュー中に選択状態が変わっても同じ対象を更新
- フォントサイズは環境設定「文字の単位」に従い、小数第2位まで換算表示
- 日本語／英語インターフェース対応

### 更新履歴

- v1.0.0 (20250509) : 初期バージョン
- v1.1.0 (20260428) : 変換形式を拡張し、実フォント情報プレビュー、キー操作、OK時の再適用を追加
- v1.1.1 (20260608) : 文字の部分選択（TextRange）からも実行可能に、OKをデフォルトボタン化、内部リファクタ
- v1.1.2 (20260611) : OK確定時の二重適用を廃止しクラッシュを回避（プレビュー状態をそのまま確定）

---

### スクリプト情報

- バージョン: v1.1.2
