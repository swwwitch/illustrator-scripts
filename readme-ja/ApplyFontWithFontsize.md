# ApplyFontWithFontsize

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyFontWithFontsize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ApplyFontWithFontsize.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyFontWithFontsize.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストフレームの各行を「フォント名＋サイズ（＋行送り）」の指定とみなして読み取り、行単位で適用します。

「ヒラギノ角ゴシック W3 12pt↓16pt」のように、フォント名のうしろにサイズ、さらにうしろに行送りを併記する書式です。

### 主な機能

- フォント名部分だけで検索してフォントを当て、併記されたサイズ・行送りも合わせて適用
- グループ内のテキストフレームも再帰的に対象（ロック・非表示は除外）
- PostScript名、ファミリー名＋スタイル名、ファミリー名による照合

### 使い方

1. 各行に「フォント名 サイズ↓行送り」を書いたテキストフレームを用意します。
2. そのテキストフレームを選択します。
3. スクリプトを実行します。

### 注意点

- サイズ・行送りを併記しない場合は ApplyFontByLine.jsx と同じ動作になります。
- 該当するフォントが見つからない行はそのまま残ります。

### 更新履歴

- v1.3.3
