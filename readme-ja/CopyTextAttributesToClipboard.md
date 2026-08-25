# CopyTextAttributesToClipboard

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyTextAttributesToClipboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/CopyTextAttributesToClipboard.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyTextAttributesToClipboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストの先頭文字を基準に、文字属性・段落属性を取得して保存します。

保存先は永続エンジン "FontClipboard" の `$.global.FontClipboard` で、ApplyTextAttributesFromClipboard.jsx から読み取って適用できます。

### 主な機能

- フォント（PostScript名・ファミリ名・スタイル）／フォントサイズ／行送り／自動行送り
- 文字ツメ／トラッキング／自動カーニング（種別と表示名）／プロポーショナルメトリクス
- 組み方向（値と表示名）／行揃え（値と表示名）

### 使い方

1. コピー元のテキストを選択します（文字ツールで部分選択しても構いません）。
2. スクリプトを実行します。
3. ApplyTextAttributesFromClipboard.jsx で適用します。

### 注意点

- 取得の基準になるのは、選択範囲の先頭文字です。
- 保存内容は `#targetengine "FontClipboard"` に置かれるため、Illustratorを終了すると失われます。

### 更新履歴

- v1.3.0 (2026-05-21)
