# 単位と数値インクリメントを変更

[![Direct](https://img.shields.io/badge/Direct%20Link-PreferenceManager--print--pt.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/single-function/PreferenceManager-print-pt.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-print-pt.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

単位と数値インクリメントを、ダイアログから指定・変更します。

### 主な機能

- 現在の単位設定を読み取って初期値に反映
- 単位と数値インクリメントをまとめて変更
- Illustratorのバージョンに応じた内部マッピングで、GUIからの補正を不要に

### 使い方

1. スクリプトを実行します。
2. 単位と数値インクリメントを指定して［OK］をクリックします。

### 注意点

- AppleScript や Keyboard Maestro を使わずに、直接設定できます。

### 更新履歴

- v1.0 (2025-08-06)
- v1.1.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (2026-09-28) : 英語表示の項目名のコロンの後ろの空白をなくし、日本語表示と同じ形にそろえた
- v1.1.2 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.1.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.1.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.5（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.6（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.7（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
