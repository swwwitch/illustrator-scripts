# 同じシンボルのインスタンスをすべて選択

[![Direct](https://img.shields.io/badge/Direct%20Link-FindAllSymbolInstances.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/FindAllSymbolInstances.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FindAllSymbolInstances.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択中のシンボルインスタンスと同じシンボルのインスタンスを、ドキュメント全体から探してまとめて選択し直します。

### 主な機能

- グループ内にネストされたシンボルインスタンスも再帰的に収集
- 複数のシンボルが混在した選択にも対応（シンボルごとに検索し、結果をまとめて選択）
- シンボルが1つも含まれない選択では、［選択］メニュー→［共通］→［アピアランス］を実行

### 使い方

1. 基準にしたいシンボルインスタンスを選択します（インスタンスを含むグループごと選択しても構いません）。
2. スクリプトを実行します。

### 注意点

- シンボルごとに最初の1つを代表として選び、［選択］メニュー→［共通］→［シンボルインスタンス］を実行した結果をまとめて選択します。
- ロックなどで選択できないアイテムはスキップします。
- 1つも選択できなかった場合はアラートを表示します。

### 紹介記事

https://note.com/dtp_tranist/n/n140952ad5011

### 更新履歴

- v1.1.1 (2026-09-15) 紹介記事へのリンクを追加。アラートを日本語／英語に対応し、内部の命名と処理を整理
- v1.1.0 (2026-05-09)
