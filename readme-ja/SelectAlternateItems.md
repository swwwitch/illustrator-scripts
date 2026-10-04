# 奇数番目／偶数番目だけを互い違いに選択

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectAlternateItems.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectAlternateItems.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectAlternateItems.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択中のオブジェクトを並び順で数え、奇数番目または偶数番目だけを互い違いに選択し直します。

### 主な機能

- ［選択］で「奇数番目」「偶数番目」を切り替え
- ［数える順］で「垂直」「水平」「重ね順」を切り替え
- ダイアログ内でのプレビュー
- 日本語／英語UI

### 使い方

1. 対象のオブジェクトを複数選択します。
2. スクリプトを実行します。
3. ［選択］と［数える順］を指定します。
4. ［OK］で確定します。

### 注意点

- ドキュメントが開いていない、またはオブジェクトを選択していないときは、警告を表示して終了します。
- ［数える順］が「垂直」「水平」のときは位置で、「重ね順」のときは重ね順で数えます。

### 紹介記事

https://note.com/dtp_tranist/n/nbad562738e70

### 更新履歴

- v1.1.0
- v1.1.2 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.3 (2026-09-29) : キーボードショートカットを共通の部品にした（⌘などを押しているときは反応しない）
- v1.1.4 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.1.5 (2026-09-29) : 前回の［数える順］が保存されていなかった不具合を修正。［方向］を［数える順］に改名し、ボタンを中央に配置。重ね順を選択の並びで数えるようにした（グループやレイヤーをまたいでも正しく数える）
- v1.1.6 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.7（2026-10-01）ボタン行を共通部品にそろえた。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.8（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.9（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
