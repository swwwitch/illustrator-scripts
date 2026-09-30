# テキストフレームを行・列単位で整列またはグループ化

[![Direct](https://img.shields.io/badge/Direct%20Link-TextGridAligner.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/TextGridAligner.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextGridAligner.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- Illustratorでテキストフレームを行・列単位で整列またはグループ化するスクリプト
- 行方向と列方向のしきい値を独立して調整可能

### 主な機能

- 行方向：天地中央に整列、またはグループ化
- 列方向：左右中央に整列、またはグループ化
- 行・列のアキを均等に配置するオプション

### 処理の流れ

1. ダイアログを表示して設定を取得
2. 選択したテキストフレームを対象に処理
3. 行・列ごとに整列またはグループ化を実行

### 更新履歴

- v1.0 (20250802) : 初期バージョン
- v1.1 (20250803) : ダイアログUI改善とコード整理
- v1.1.2 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.3 (20260928) : ボタン行を共通の部品で組むようにした
- v1.1.4 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.5 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.6 (20260930) : 初めて開くときにダイアログを右へずらす独自の配置をやめた
- v1.1.7（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.8（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた

### スクリプト情報

- バージョン: v1.1.6
