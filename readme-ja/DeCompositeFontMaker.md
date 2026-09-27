# 合成フォントを構成フォントに置き換え

[![Direct](https://img.shields.io/badge/Direct%20Link-DeCompositeFontMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/DeCompositeFontMaker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeCompositeFontMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストに使われている合成フォントを、文字ごとに構成フォント（和文・かな・欧文などのフォント）へ置き換えます。合成フォントで指定したサイズ・比率・ベースラインも文字の書式に写すので、見た目を保ったまま通常のフォントに戻せます。合成フォントを持っていない環境に渡すデータの準備などに使えます。

Illustrator のスクリプトからは合成フォントの中身を取り出せないため、合成フォントのフォルダーにあるファイルを直接読みます。

### 主な機能

- 文字ごとに、漢字・かな・全角約物・全角記号・半角欧文・半角数字のどの文字セットにあたるかを判定し、その文字セットのフォントを当てる
- 合成フォントの設定を文字の書式に写す
  - **サイズ**：文字サイズに掛ける（例：24pt・欧文116% → 27.84pt）
  - **垂直比率・水平比率**：文字の比率に掛ける
  - **ベースライン**：元の文字サイズに対する比率をベースラインシフトに足す（例：24pt・−2% → −0.48pt）
- 自動行送りの文字は、置き換え前の行送りを固定値で残す（サイズの違う文字セットがある合成フォントのとき）
- トラッキングと手動カーニングは、置き換え後の文字サイズに合わせて換算する（1/1000em 単位のため）
- 合成フォントの文字だけを置き換え、通常のフォントの文字はそのまま
- 置き換えた文字数と、置き換えられなかった文字の理由をアラートで知らせる

### 使い方

1. テキストオブジェクトを選択します（グループの中のテキスト、文字の選択でも可）
2. スクリプトを実行します

### 置き換えられないとき

| 状況 | 動作 |
|---|---|
| 合成フォントのファイルがこの環境に無い | その合成フォントの文字はそのまま。ドキュメントに記録された構成フォント名（ファイル名）を知らせる |
| 構成フォントがこの環境に無い | その文字はそのまま。フォント名と文字数を知らせる |

### 注意点

- 特例文字セットは無視します。特例文字セットの文字は、標準の文字セット（漢字・かな・欧文など）の設定で置き換えます
- 合成フォントのファイルは `~/Library/Application Support/Adobe/Adobe Illustrator <バージョン>/ja_JP/合成フォント/` から探します
- 行送りの固定と、トラッキング・カーニングの換算は、スクリプト冒頭のユーザー設定（`KEEP_LEADING` / `KEEP_SPACING`）で止められます

### 更新履歴

- v1.0.0 (20260927) : 初版リリース
