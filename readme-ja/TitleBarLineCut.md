# テキストの周りで罫線が欠けたタイトル帯を作成

[![Direct](https://img.shields.io/badge/Direct%20Link-TitleBarLineCut.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/TitleBarLineCut.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TitleBarLineCut.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

テキスト1つと長方形パス1つを選択して実行すると、テキストの周りで罫線が欠けたタイトル帯を作成します。

### 主な機能

- マージン、角丸、塗り、ノッチ、線幅をダイアログで指定
- 日本語／英語UI

### 使い方

1. テキスト1つと長方形パス1つを選択します。
2. スクリプトを実行します。
3. マージンなどを指定して［OK］をクリックします。

### 注意点

- マージンは0以上の数値で指定します。
- ダイアログの値はセッション内で保持され、Illustratorを再起動するとリセットされます。
- 帯を使わず罫線だけを欠けさせたい場合は TitleBarLineCutSp.jsx を使用してください。

### 更新履歴

- v1.2.7（2026-10-01）初回にダイアログを右へずらす独自の配置をやめ、位置を共通部品に一本化。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.2.6（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.2.5 (2026-09-30) 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.2.4 (2026-09-30) 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.2.3 (2026-09-29) ダイアログの不透明度を98%に変更
- v1.2.2 (2026-09-28) ボタン行を共通の部品で組むようにした
- v1.2.1 (2026-09-28) ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.2.0 (2026-09-27) 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.2
