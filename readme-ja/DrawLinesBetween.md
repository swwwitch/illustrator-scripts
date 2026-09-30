# オブジェクトの間に水平の罫線を描く

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawLinesBetween.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/DrawLinesBetween.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawLinesBetween.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクト（図形／テキスト）を上から順に並べ、その間に水平の罫線を描画します。

### 主な機能

- 入力単位は環境設定の「線」（`strokeUnits`）に追従（表示ラベルと内部のpt換算の両方）
- ［延長］で罫線を左右方向に伸縮（＋で延長、−で短縮）

### 使い方

1. 罫線を挟みたいオブジェクトをまとめて選択します。
2. スクリプトを実行します。
3. 線幅や延長量を指定して実行します。

### 更新履歴

- v1.1.7（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.6（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.1.5 (20260930) : 初めて開くときにダイアログを右へずらす独自の配置をやめた。右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.1.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした。線幅・延長・線端の前回値が保存されていなかった不具合を修正（Illustrator に無い Photoshop 用の関数を使っていた。保存先: `~/Library/Application Support/illustrator-scripts/DrawLinesBetween.json`）
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0
