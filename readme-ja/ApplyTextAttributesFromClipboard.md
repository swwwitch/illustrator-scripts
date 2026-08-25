# ApplyTextAttributesFromClipboard

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyTextAttributesFromClipboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ApplyTextAttributesFromClipboard.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyTextAttributesFromClipboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

CopyTextAttributesToClipboard.jsx が保存した文字属性を、選択中のテキストへ適用します。

永続エンジン "FontClipboard" の `$.global.FontClipboard` から読み取るため、2つのスクリプトを組み合わせて「書式のコピー＆ペースト」として使います。

### 主な機能

- フォント・サイズ／行送り、カーニング関連、段落属性、塗りとグラフィックスタイルの4パネル構成
- 各属性はチェックボックスで適用可否を切り替え（初期状態でONなのはフォントのみ）
- 「塗りとグラフィックスタイル」パネルはラジオで排他（しない／塗り／グラフィックスタイル）

### 使い方

1. CopyTextAttributesToClipboard.jsx で属性をコピーしておきます。
2. 適用先のテキストを選択します（文字ツールで部分選択しても構いません）。
3. スクリプトを実行し、適用する属性にチェックを入れて［OK］をクリックします。

### 注意点

- テキスト編集モードで部分選択している場合は、その範囲だけに適用します。
- 複数の TextFrame を選択している場合は、それぞれに同じ属性を適用します。
- 属性の受け渡しには `#targetengine "FontClipboard"` を使うため、Illustratorを終了すると内容は失われます。

### 更新履歴

- v1.3.1 (2026-05-21)
