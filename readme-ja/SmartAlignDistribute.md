# 縦横を自動判定して指定間隔で分布

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAlignDistribute.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignDistribute.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトを「縦/横」に並べて、指定した間隔で分布します。方向は自動判定も可能で、揃え（左右/上下）、プレビュー境界（visible/geometric）、ランダム並べ替えにも対応します。
縦並び時にテキストを含む場合、ダイアログ中だけ一度だけ計測用に複製→アウトライン化して高さ（必要に応じて幅も）を計測し、その結果をダイアログ中だけキャッシュします（プレビューのたびに複製しない）。

### オリジナルアイデア

John Wundes
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

### スクリプト情報

- バージョン: v1.3.7
- 最終更新: 2026-10-04

### 更新履歴

- v1.3.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.3.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.2 (20260928) : ボタン行を共通の部品で組むようにした。キーボードショートカットを共通の部品にした（⌘などを押しているときは反応しない）。クリップグループの範囲をマスクで測るようにした（テキストのマスク、グループの中のクリップグループにも対応）
- v1.3.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.3.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.3.5（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.3.6（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.3.7（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
