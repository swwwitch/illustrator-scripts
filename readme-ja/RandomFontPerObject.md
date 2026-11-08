# オブジェクトごとにフォントをランダムに適用

[![Direct](https://img.shields.io/badge/Direct%20Link-RandomFontPerObject.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/RandomFontPerObject.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RandomFontPerObject.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストオブジェクトごとに、フォントをランダムに適用します。1つのオブジェクトの中は同じフォントにそろいます。[オート文字タッチツール](AutoTouchType.md)のフォントの入れ替えを、1文字ごとからオブジェクトごとに変えて切り出したスクリプトです。同じ機能は、[位置・スケール・回転などをまとめてランダムに変化させるスクリプト](RandomizeObjects.md)にもあります。

### 主な機能

- オブジェクトごとに1つのフォントを選んで適用します。グループの中のテキストも対象です。
- ［和文フォントに限定］をオンにすると、候補を和文フォントだけに絞ります。和文を含むテキストを選択しているときは、最初からオンになります。
- ［ランダム］で、同じ設定のまま別のフォントを引き直します。

### 使い方

1. テキストオブジェクトを選択します。
2. スクリプトを実行します。開いた時点でフォントが変わります。
3. ［ランダム］や［和文フォントに限定］で調整します。
4. ［OK］で確定します。［キャンセル］で元に戻ります。

### 注意点

- 文字ツールで文字を選択しているときは、その範囲を1つのオブジェクトとして扱います。
- ロックされたテキストと非表示のテキストは対象外です。
- 和文フォントかどうかは、フォント名から判定しています。

### 更新履歴

- v1.0.0 (2026-11-07) : 初版
- v1.0.1 (2026-11-08) : 和文フォントの判定で、フォントワークスの中国語版・韓国語版、FWArabic などの和文以外の書体、BM Dohyeon・배달의민족 도현 などの韓国語フォント、フォントワークス絵文字、漢字を持たないかなフォント、中国語のフォント（PingFang・儷黑・Hiragino Sans GB など）、靜呉竹、Tsukimi、名前に Thai・體 を含むフォントを和文として扱わないようにした
- v1.0.2 (2026-11-08) : 和文フォントの判定で、HGで始まるフォントを和文として扱わないようにした
