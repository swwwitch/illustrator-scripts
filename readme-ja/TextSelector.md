# TextSelector

[![Direct](https://img.shields.io/badge/Direct%20Link-TextSelector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextSelector.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ドキュメント内のテキストフレームを、複数の条件で一括選択します。

### 主な機能

- 属性で選択: 選択中テキストを基準に、フォントファミリー／＋スタイル／＋サイズ／フォントサイズ／テキストカラー／不透明度で検索
- テキストの種類: すべて／ポイント文字／エリア内文字／パス上文字
- 文字列で選択: 完全一致／部分一致／先頭一致／末尾一致／正規表現
- 選択後の処理: なし／非表示／「_text」レイヤーへ移動／一括編集

### 使い方

1. （属性で選択する場合は）基準にするテキストを選択します。
2. スクリプトを実行します。
3. 条件と選択後の処理を指定して実行します。

### 注意点

- 「_text」レイヤーへ移動する場合、既存レイヤーのロック・可視状態は復元します。
- 一括編集では、書式（`characterAttributes`）を維持したまま内容を置換します。

### 更新履歴

- v1.2.5
