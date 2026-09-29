# 重ね順を位置やZインデックスで並べ替え

[![Direct](https://img.shields.io/badge/Direct%20Link-ZIndexSorter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/sort/ZIndexSorter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZIndexSorter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- オブジェクトの重ね順を位置（X/Y）やZインデックスで並べ替える
- 昇順／降順／ランダムに並べ替え可能

### 主な機能

- ダイアログで並べ替え軸と順序を指定
- キャンセル時には元の順序に復元
- 多言語対応（日本語／英語）

### 処理の流れ

- 選択アイテムを収集
- 並べ替え軸・順序の設定をダイアログで取得
- 並べ替え後、Zインデックス順に再配置

### 注意点

- Illustrator 2025 以降で動作確認済

### 更新履歴

- v1.0 (20250806) : 初期バージョン
- v1.0.2 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.0.3 (20260928) : ボタン行を共通の部品で組むようにした
- v1.0.4 (20260929) : ダイアログの不透明度を98%に変更
- v1.0.5 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正

### スクリプト情報

- バージョン: v1.0.4
