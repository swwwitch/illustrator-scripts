# 重なったオブジェクトを横方向に等間隔で並べ直す

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAlignAndTile--simple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/single-function/SmartAlignAndTile-simple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile-simple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

重なって配置されたオブジェクトを、横方向へ等間隔に並べ直します。SmartAlignAndTile.jsx の簡易版です。

### 使い方

1. 対象のオブジェクトを選択します。
2. スクリプトを実行します。

### 注意点

- 常駐エンジン（`#targetengine`）で動作します。
- 元アイデア: John Wundes「Distribute Stacked Objects v1.1」（[js4ai](https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx)）、[Gorolib Design](https://gorolib.blog.jp/archives/77282974.html)

### 更新履歴

- v1.1.7（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.1.6（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.5（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした。キーボードショートカットを共通の部品にした（⌘などを押しているときは反応しない）。クリップグループの範囲をマスクで測るようにした
- v1.0.2
