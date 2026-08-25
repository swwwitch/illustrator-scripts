# KeepInView

[![Direct](https://img.shields.io/badge/Direct%20Link-KeepInView.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/_templates/KeepInView.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KeepInView.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

生成・変更したオブジェクトが画面から外れたときだけ、見える位置へ表示を移す再利用テンプレートです。

すでに見えているときは動かさないので、操作のたびに画面が揺れません。

### 主な機能

- 対象が画面外に出たときだけスクロールします
- 画面に収まらないときはズームアウトのみ行い、拡大はしません（倍率が勝手に上がらない）
- 日本語／英語のラベルとツールチップを内蔵したチェックボックスを追加できます

### 使い方

1. `var KeepInView = (function () { ... })();` のブロックまるごとを、対象スクリプトのIIFE内へコピーします。
2. チェックボックスを置きます。文言は内蔵しているため、別途用意する必要はありません。

        var cbKeepInView = KeepInView.addCheckbox(grpViewOptions, { value: true });

3. 結果を作り終えたところで呼びます。

        if (cbKeepInView.value) KeepInView.ensureVisible(createdItems, { doc: doc, fitRatio: 0.9 });

### 注意点

- チェックボックスを自前で作る場合は `addCheckbox` を使わず、`ensureVisible` だけを呼びます。
- `doc` を省略すると最前面のドキュメント、`fitRatio` を省略すると `DEFAULT_FIT_RATIO` が使われます。
- `#include` は使わず、各スクリプトを1ファイルで完結させる方針のためコピーして使います。
- 使用例: `jsx/text/DynamicTextGenerator.jsx`（「結果を画面内に表示」）

### 更新履歴

- v1.0.0: 初版（テンプレート）
