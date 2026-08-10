# InspectKinsokuSimple

[![Direct](https://img.shields.io/badge/Direct%20Link-InspectKinsokuSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/InspectKinsokuSimple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectKinsokuSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したテキストフレームの各段落で使われている禁則処理セットの値を集め、一覧をアラートで表示します。
- ドキュメント内で禁則設定が混在していないかを確認するための小さな調査用スクリプトです。

### 使い方

1. 調べたいテキストフレームを選択する
2. スクリプトを実行する
3. 検出された禁則値がアラートに一覧表示される

### 注意点

- 禁則「なし」は属性が取得できずエラー（9563）になるため、「なし」として扱います。
- ドキュメントの内容は変更しません。
