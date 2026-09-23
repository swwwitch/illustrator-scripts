# 混在した文字サイズを先頭文字のサイズに統一

[![Direct](https://img.shields.io/badge/Direct%20Link-FontSizeToScaleConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/FontSizeToScaleConverter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontSizeToScaleConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

1つのテキストの中で、一部の文字だけ文字サイズを変えていることがあります。見出しの中の数字を大きくする、単位を小さくする、といった場合です。こうしたテキストは、あとで全体の文字サイズを変えたり、文字スタイルを当てたりするときに扱いにくくなります。

このスクリプトは、テキスト内の文字サイズを先頭文字のサイズにそろえます。サイズの差は水平比率・垂直比率に置き換えるので、見た目の大きさは変わりません。

たとえば先頭が 10pt のテキストに 15pt の文字があれば、その文字は 10pt・水平比率150%・垂直比率150% になります。

### 主な機能

- テキストごとに、先頭文字の文字サイズを基準にして、ほかの文字のサイズをそろえる
- サイズの差を水平比率・垂直比率に置き換え、見た目の大きさを保つ。すでに比率が付いている文字は、その比率に掛け合わせる
- 自動行送りの文字は、変換前の行送りを固定値にして、行間が変わらないようにする
- ポイント文字・エリア内文字・パス上文字に対応
- 複数のテキストやグループ内のテキストもまとめて処理

### 使い方

1. 文字サイズが混在しているテキストオブジェクトを選択します（複数選択やグループも可）
2. スクリプトを実行します

ダイアログは出ず、すぐに変換されます。

### オプション

スクリプト冒頭のユーザー設定で、次の項目を変えられます。

| 設定 | 内容 |
|---|---|
| `PRESERVE_LEADING` | `false` にすると、自動行送りの文字の行送りを固定しない（初期値 `true`） |
| `FONT_SIZE_TOLERANCE` | 同じサイズとみなす差の許容値（pt、初期値 `0.001`） |

### 注意点

- 基準はテキストごとの先頭文字です。複数のテキストを選んでも、テキストどうしでサイズはそろえません
- 変換後の比率が 1%〜10000% の範囲を超える文字は、見た目を保てないので変換しません
- 自動行送りの文字は、行送りが固定値に変わります。あとで文字サイズを変えても、行送りは追従しません
- ロックまたは非表示のオブジェクトは対象外です

### 更新履歴

- v1.0.0（2026-06-18）：初期バージョン
