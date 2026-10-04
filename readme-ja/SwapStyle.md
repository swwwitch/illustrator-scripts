# 2つのオブジェクトの見た目または文字列を交換

[![Direct](https://img.shields.io/badge/Direct%20Link-SwapStyle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/SwapStyle.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapStyle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した2つのオブジェクトの間で、見た目または文字列を交換します。

### 主な機能

- ダイアログで「スタイル交換」「文字列交換」を切り替え
- スタイル交換は次の3系統を組み合わせて指定
  - グラフィックスタイル交換（現在のアピアランス全体）
  - 基本的な塗りや線（塗り／線のカラー／線幅）
  - 文字属性

### 使い方

1. 交換したい2つのオブジェクトを選択します。
2. スクリプトを実行します。
3. 交換する内容を指定して［OK］をクリックします。

### 注意点

- 選択が2つでない場合は処理を行いません。

### 更新履歴

- v1.1.0 (2026-05-23)
- v1.1.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (2026-09-28) : 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした。ボタン行を共通の部品で組むようにした
- v1.1.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.1.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.5（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.6（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.7（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
