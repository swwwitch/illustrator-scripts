# ［ぼかし］のライブエフェクトを適用

[![Direct](https://img.shields.io/badge/Direct%20Link-LEFeather.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/single-function/LEFeather.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LEFeather.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトに［効果］＞［スタイライズ］＞［ぼかし］のライブエフェクトを適用します。

### 使い方

1. ぼかしたいオブジェクトを選択します。
2. スクリプトを実行します。
3. 半径を入力して［OK］をクリックします。

### 注意点

- 半径は現在の定規単位（`rulerType`）で表示・入力し、内部では pt に換算します。

### 更新履歴

- v1.0.0
- v1.1.0 (2026-09-27): 半径の入力欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ、option＋で0.1ずつ）
- v1.1.1 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (2026-09-28): 半径の単位表示を修正（定規単位が歯のとき「Q」ではなく「H」、フィート・メートル・ヤードが pt 扱いだった不具合も修正）。インチの表記を「in」に統一。ボタン行を共通の部品で組むようにした
- v1.1.3 (2026-09-29): ダイアログの不透明度を98%に変更
- v1.1.4 (2026-09-30): 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.5 (2026-09-30): 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.1.6（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.1.7（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.8（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
