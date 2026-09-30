# 指定した文字の直後で改行する

[![Direct](https://img.shields.io/badge/Direct%20Link-titlemaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/titlemaker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/titlemaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択中のテキストフレームを、指定した文字の直後で改行する
- 改行対象の文字（、。〜）をチェックボックスで選択
- 改行後、連続する改行を1つにまとめる
- 文字揃えを欧文ベースラインに設定

### 処理の流れ

1. ダイアログで改行対象文字を選択
2. 選択中の各テキストフレームの内容を走査
3. 対象文字の後ろに改行を挿入し、重複改行を整理
4. 欧文ベースラインを適用

### スクリプト情報

- バージョン: v1.1.2

### 更新履歴

- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.3 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
