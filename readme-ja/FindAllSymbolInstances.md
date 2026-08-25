# FindAllSymbolInstances

[![Direct](https://img.shields.io/badge/Direct%20Link-FindAllSymbolInstances.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/FindAllSymbolInstances.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FindAllSymbolInstances.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択中のシンボルインスタンスと同じシンボルのインスタンスを、ドキュメント全体から探してまとめて選択し直します。

### 主な機能

- グループ内にネストされたシンボルインスタンスも再帰的に収集
- 複数のシンボルが混在した選択にも対応
- シンボルが1つも含まれない選択では、同一アピアランスの検索にフォールバック

### 使い方

1. 基準にしたいシンボルインスタンスを選択します。
2. スクリプトを実行します。

### 注意点

- シンボル定義ごとに最初の1つを代表として選び、［シンボルインスタンスを選択］コマンドを実行した結果を重複なくマージします。
- ロックなどで再選択できないアイテムは黙ってスキップします。
- 何も集約できなかった場合のみアラートを表示します。

### 更新履歴

- v1.1.0 (2026-05-09)
