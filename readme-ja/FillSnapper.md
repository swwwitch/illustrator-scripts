# 対象のバウンディングボックスを最寄りの基準線に合わせる

[![Direct](https://img.shields.io/badge/Direct%20Link-FillSnapper.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/FillSnapper.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillSnapper.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択中のオブジェクトを「動かす対象」と「スナップ基準」に分類し、対象のバウンディングボックスを最寄りの基準線へ合わせます。

パスはアンカーポイントを直接変形するため、クリップグループ内の子パスにも対応します。

### 使い方

1. 動かしたいオブジェクトと、基準にする罫線をまとめて選択します。
2. スクリプトを実行します。
3. 許容差とスナップ距離を指定して実行します。

### オプション

**線判定の許容差**

細長いパスを水平線／垂直線として扱うための判定幅です。

**最大スナップ距離**

対象の辺からどれだけ離れた基準線まで吸着するかの制限です。0 の場合は距離制限なしになります。

### 更新履歴

- v1.0.1
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
