# SmartSelectionFilterSimple

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSelectionFilterSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SmartSelectionFilterSimple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSelectionFilterSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトを条件に応じてフィルタリングし、テキスト／オープンパス／クローズパスを選択し直します。

対象スコープを切り替えると、選択直下だけでなくグループ内のオブジェクトも対象にできます。

### 主な機能

- テキスト／オープンパス／クローズパスでの絞り込み
- 対象スコープの切り替え（選択直下のみ／グループ内も含める）
- クローズパスは塗りのみ／線のみ／塗り＋線のいずれも対象

### 使い方

1. 対象のオブジェクトを選択します。
2. スクリプトを実行します。
3. 条件と対象スコープを指定して確定します。

### 注意点

- 複合パスは親オブジェクトとして扱います。
- クリップグループでは、マスク用パスを選択対象から除外します。
- 選択の更新時は、ロック／非表示のオブジェクトや親階層を避けて処理します。
- より詳しい条件を指定したい場合は SmartSelectionFilter.jsx を使用してください。

### 更新履歴

- v1.0
