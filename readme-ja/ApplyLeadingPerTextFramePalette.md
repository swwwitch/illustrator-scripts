# 行頭の文字サイズを基準に行送りを再計算

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyLeadingPerTextFramePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ApplyLeadingPerTextFramePalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyLeadingPerTextFramePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択された各テキストフレームの各行について、行頭数文字のフォントサイズを基準に行送りを再計算して適用します。

適用する行送りの割合はダイアログで指定できます。

### 使い方

1. 対象のテキストフレームを選択します。
2. スクリプトを実行します。
3. 行送りの割合を指定して実行します。

### 注意点

- テキストの一部（TextRange）を選択している場合は、その親のテキストフレームに正規化して処理します。
- 行送りの値そのものではなく、自動行送りの値（％）を変更して調整します。
- 割合を固定した派生版として ApplyLeadingPerTextFrame110.jsx / 150.jsx / AUTO.jsx があります。

### 更新履歴

- v1.2.2（2026-10-03）常駐エンジンを共通エンジン `SwwwitchPalettes` に変更し、［常駐パレットをまとめて閉じる］で閉じられるようにした
- v1.2.1（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.2.0（2026-09-27）数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）。
- v1.1.2（2026-09-26）ファイル名を `ApplyLeadingPerTextFrame.jsx` から `ApplyLeadingPerTextFramePalette.jsx` に変更。
- v1.1.0 (2026-07-08)
