# クリップ形状を正方形・正円・六角形に置き換え

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipMaskShapeChanger.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipMaskShapeChanger.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipMaskShapeChanger.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した画像（配置／埋め込み）または既存のクリップグループを、正方形・正円・六角形（A/B）のクリップ形状へ置き換えます.

### 主な機能

- クリップ形状を正方形／正円／六角形A／六角形Bから選択
- ［ケイ線を追加］: 生成後にケイ線を追加
- ［角丸］: 生成したクリップグループに角丸のライブエフェクトを適用（正円では無効）
- ［複数オブジェクト＞大きさを揃える］: 選択内の最大／最小（面積）を基準にサイズをそろえる

### 使い方

1. 対象の画像、または既存のクリップグループを選択します。
2. スクリプトを実行します。
3. 形状とオプションを指定して［OK］をクリックします。

### 更新履歴

- v1.0 (2026-02-01)
- v1.1.0 (2026-09-27): 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (2026-09-28): ボタン行を共通の部品で組むようにした
- v1.1.3 (2026-09-29): ダイアログの不透明度を98%に変更
