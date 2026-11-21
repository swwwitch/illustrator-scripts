# ドキュメント全体を CMYK のモノクロにする（テキストは K100）

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ConvertDocumentToGrayscale.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- ドキュメントを CMYK にしてから、すべてのオブジェクトをグレースケールに変換します。モノクロ印刷用のデータ作りに使います。
- テキストは K100 にし、最後にスウォッチを削除します。

### 主な機能

次の順に処理します。

1. ［編集］→［カラーを編集］→［CMYK に変換］で、オブジェクトの色を CMYK にする
2. RGB のドキュメントは、［ファイル］→［ドキュメントのカラーモード］→［CMYK カラー］に切り替える
3. シンボルの中身をグレースケールに変換する（入れ子のシンボルも対象）
4. ［編集］→［カラーを編集］→［グレースケールに変換］で、すべてのオブジェクトをモノクロにする
5. テキストの塗りと線を K100 にする（シンボルの中のテキストも対象）
6. スウォッチをすべて削除し、空になったスウォッチグループも片付ける

ロック・非表示のレイヤーとオブジェクトも一時的に解除して処理し、終わったら元に戻します。

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する

### 注意点

- ダイアログボックスはありません。選択の有無にかかわらず、ドキュメント全体が対象です。
- テキストで K100 にするのは、塗りと線の単色だけです。白（グレー 0%）は白のまま残します（スクリプト冒頭の KEEP_WHITE_TEXT を false にすると白も K100）。グラデーション・パターンの文字はグレースケールのままです。
- シンボルは、中身をグレースケールにした新しいシンボルにインスタンスを差し替え、元のシンボルと入れ替えます。基準点、ダイナミックシンボルの種類、9スライスなどのシンボルオプションは引き継がれません。
- リンク画像、パターン・ブラシの中身、グラフィックスタイルは変換されません。
- ［なし］［レジストレーション］など削除できないスウォッチは残ります。
- 取り消すときは、［取り消し］を何度か実行します。

### 更新履歴

- v1.0.0 (20261121) : 初版
- v1.1.0 (20261121) : ［CMYK に変換］とドキュメントのカラーモードの CMYK への切り替え、シンボルの中身のグレースケール変換、テキストの K100 を追加
