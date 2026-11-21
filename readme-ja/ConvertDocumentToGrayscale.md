# ドキュメント全体をグレースケールに変換してスウォッチを削除

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ConvertDocumentToGrayscale.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- ドキュメント内のすべてのオブジェクトを選択し、［編集］→［カラーを編集］→［グレースケールに変換］でモノクロにします。
- 続けてスウォッチをすべて削除し、空になったスウォッチグループも片付けます。

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する

### 注意点

- ダイアログボックスはありません。選択の有無にかかわらず、ドキュメント全体が対象です。
- ロック・非表示のレイヤーとオブジェクトも一時的に解除して変換し、終わったら元に戻します。
- リンク画像、シンボルの定義、パターン・ブラシの中身、グラフィックスタイルは変換されません。
- ［なし］［レジストレーション］など削除できないスウォッチは残ります。
- 取り消すときは、［取り消し］を何度か実行します。

### 更新履歴

- v1.0.0 (20261121) : 初版
