# 複合パスを解除して合体・拡張しオフセット

[![Direct](https://img.shields.io/badge/Direct%20Link-PathUniteOffsetTool.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/PathUniteOffsetTool.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathUniteOffsetTool.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択中のオブジェクトに対して複合パス解除 → パスの合体 → アピアランス拡張 → グループ解除 → オフセットパスを一括実行
- ダイアログでプレビューを有効化すると、閉じずに結果を確認可能

### 主な機能

- 複合パスの解除
- パスの合体（ライブパスファインダ）
- アピアランスの拡張
- グループ解除
- 指定値（mm）でオフセットパス効果を適用
- プレビュー機能

### 処理の流れ

1) 複合パスの解除
2) パスの合体（ライブパスファインダ）
3) アピアランスの拡張
4) グループ解除
5) 指定値（mm）でオフセットパス効果を適用

### 更新履歴

- v1.0.0 (2026-05-10) : 初期バージョン
- v1.1.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (2026-09-28) : オフセットの単位表示を修正（定規単位が歯のとき「Q」ではなく「H」、フィート・メートル・ヤードが pt 扱いだった不具合も修正）。インチの表記を「in」に統一
- v1.1.2 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.1.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.1.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正

### スクリプト情報

- バージョン: v1.1.3
