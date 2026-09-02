# カーソルのある段落をひとつ下へ移動

[![Direct](https://img.shields.io/badge/Direct%20Link-moveParagraphDown.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/moveParagraphDown.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/moveParagraphDown.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 文字ツールでカーソルを置いた段落を、ひとつ下の段落とまるごと入れ替えます。Visual Studio Code の「行を下へ移動」を、表示行ではなく段落単位で行うイメージです。
- sky-chaser-high 氏の moveLineDown.jsx を、段落単位で動くように改変したものです。
- 入れ替えたあともカーソルは移動した段落の中に残るので、続けて実行すればその分だけ下へ送れます。

### 主な機能

- 段落全体を選択する必要はなく、段落内にカーソルを置くだけで実行できる
- フォント・サイズ・段落スタイルなどの書式ごと入れ替える
- 実行前と同じ段落内の位置にカーソルを戻すので、連続実行できる
- 最後の段落では何もしない

### 使い方

1. 文字ツールで、移動したい段落の中にカーソルを置きます（段落内の文字を選択した状態でも構いません）。
2. スクリプトを実行します。
3. 繰り返し実行すると、その回数だけ下へ移動します。

アクションに登録してファンクションキーを割り当てておくと、テキストエディター感覚で段落を並べ替えられます。

### 注意点

- 文字ツールでテキストを編集している（カーソルがテキスト内にある）ときだけ動作します。オブジェクトとしてテキストを選択しているだけでは何も起こりません。
- 段落の入れ替えにクリップボードを使うため、実行するとクリップボードの内容が置き換わります。
- 移動先が空段落（改行だけの段落）のときは、カーソル位置の復帰を行いません。
- 逆方向へ動かす moveParagraphUp.jsx もあります。

### 更新履歴

- v1.0.0 (2026-08-27) : 初版
- v1.0.1 (2026-09-02) : 入れ替え後のカーソル復帰で「Error 21: undefined はオブジェクトではありません。」が出る不具合を修正（Story にない `contents` を参照していたため、段落から文字列を取得する方式に変更）
