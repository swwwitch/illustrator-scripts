# 奇数番目／偶数番目だけを互い違いに選択

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectAlternateItems.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectAlternateItems.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectAlternateItems.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択中のオブジェクトを並び順で数え、奇数番目または偶数番目だけを互い違いに選択し直します。

### 主な機能

- ［選択］で「奇数」「偶数」を切り替え
- ［方向］で「垂直」「水平」「重ね順」を切り替え、数える順序を変更
- ダイアログ内でのプレビュー
- 日本語／英語UI

### 使い方

1. 対象のオブジェクトを複数選択します。
2. スクリプトを実行します。
3. ［選択］と［方向］を指定します。
4. ［OK］で確定します。

### 注意点

- ドキュメントが開いていない、またはオブジェクトを選択していないときは、警告を表示して終了します。
- ［方向］が「垂直」「水平」のときは位置で、「重ね順」のときは重ね順で数えます。

### 更新履歴

- v1.1.0
- v1.1.2 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.3 (2026-09-29) : キーボードショートカットを共通の部品にした（⌘などを押しているときは反応しない）
- v1.1.4 (2026-09-29) : ダイアログの不透明度を98%に変更
