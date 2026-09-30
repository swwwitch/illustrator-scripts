# 文字タッチツールの調整をランダムに付与

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoTouchType.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoTouchType.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoTouchType.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストの各文字に対して、ベースライン・比率・回転・カーニング・トラッキングをランダムに付与する「オート文字タッチ」ツールです。

### 主な機能

- seed付きの乱数でプレビューの見た目を安定化
- 文字回転を適用したときは、回転角に応じてトラッキングを自動補正
- ［軽量モード］で、画面ズームをドラッグ完了時のみ反映に切り替え

### 使い方

1. 対象のテキストを選択します。
2. スクリプトを実行します。
3. 各項目のばらつきを指定し、プレビューを確認して［OK］をクリックします。

### 紹介記事（note）

https://note.com/dtp_tranist/n/ne6545c4717af

### 更新履歴

- v1.3.6（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.3.5（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.3.4 (2026-09-30) 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.3.3 (2026-09-29) ダイアログの不透明度を98%に変更
- v1.3.2 (2026-09-28) クリップグループの範囲をマスクで測るようにした
- v1.3.2 (2026-09-28) ボタン行を共通の部品で組むようにした
- v1.3.1 (2026-09-28) ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.0 (2026-09-27) 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.2.9 (2026-09-19) 紹介記事へのリンクを追加。ラベル定義のカテゴリ分け、命名の見直し、重複コードの整理など内部構造を整理。
- v1.2.7 (2026-02-20)
