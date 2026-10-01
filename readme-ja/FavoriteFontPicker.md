# よく使うフォントだけを一覧から選んで適用

[![Direct](https://img.shields.io/badge/Direct%20Link-FavoriteFontPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FavoriteFontPicker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteFontPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- よく使うフォントだけに絞った一覧から選び、選択中のテキストに適用する Illustrator スクリプト
- 同じ書体の規格違い（Pr6N・Pr5 など）は、優先順位がいちばん高いものだけを一覧に残す
- KOUJI & 相棒（Gem）さんの「よく使うフォントパネル」（favoriteFont_AI.jsx v1.1.2、MIT License）をもとに改変

### 主な機能

- 左にチェックボックス、右にフォント一覧の2カラム
- ［分類］：カスタム／ドキュメント／モリサワ／TB（タイプバンク）／FOT（フォントワークス）（チェックしたものを合わせて表示）
  - カスタム：スクリプト冒頭の `CUSTOM_FONTS` に書いたフォント（前方一致）
  - ドキュメント：ドキュメントと同じフォルダーの `_ProjectFonts.txt` に書いたフォント。規格違いの間引きをせず、そのまま表示
- ［規格］：左に無印（Std・Pro・Pr5・Pr6）、右に N 付きを並べる。チェックした規格（と規格なし）だけを表示。同じ書体は、チェックした中で優先順位がいちばん高い規格を残す（Pr6N を外せば Pr6 や Pr5 が残る）
- ［すべて表示］：［分類］・規格違いの間引き・除外をやめ、すべてのフォントを表示（［規格］の絞り込みは効く）
- ［分類］［規格］のチェックボックスは、option（Alt）＋クリックで「クリックしたものだけオン」と「すべてオン」を切り替え
- ［絞り込み］：ファミリー名・スタイル名・PostScript 名の部分一致
- ［PostScript名で表示］：一覧を PostScript 名（RyuminPr6N-Light など）で表示し、その順に並べる
- ［使用フォントをフォルダーに記録］：ドキュメントで使っているフォントを `_ProjectFonts.txt` に追記（InDesign 版と同じファイル）
- ［再スキャン］：インストールされているフォントを読み直す
- チェックの状態と絞り込みの文字列は、次回も同じ状態で開く

### 使い方

1. テキストフレームまたはグループを選択する（文字ツールで一部の文字を選択してもよい）
2. スクリプトを実行する
3. 一覧のフォントをダブルクリック、または選んで Enter キーか［適用］

［絞り込み］は、結果が800件以下なら入力のたびに一覧を更新し、それより多いときは件数だけを表示して Enter キーで一覧を更新します（件数は `LIVE_SEARCH_MAX_FONTS` で変更可）。

### オプション（スクリプト冒頭のユーザー設定）

| 項目 | 内容 | 初期値 |
|---|---|---|
| CUSTOM_FONTS | ［カスタム］に出すフォント（前方一致） | Graphik, DIN |
| EXCLUDE_FONTS | 一覧から外すフォント（部分一致、［すべて表示］では外さない） | NT |
| RESCUE_FONTS | EXCLUDE_FONTS に当たっても残すフォント（部分一致） | なし |
| PRIORITY_PREFIXES | 残す接頭辞の優先順位（先頭ほど優先） | A-OTF, A P-OTF, AP-OTF, G-OTF, U-OTF |
| PRIORITY_SUFFIXES | 残す規格の優先順位。［規格］のチェックボックスにも並ぶ | Pr6N, Pr6, Pr5N, Pr5, ProN, Pro, StdN, Std |
| FOUNDRY_FILTERS | メーカー別の分類（ファミリー名・PostScript 名の前方一致） | モリサワ・TB・FOT |

### 注意点

- 初回と、フォントの構成が変わったとき（総数と、一覧から等間隔に取った32件の名前で判定）は、フォント情報を読み込んでから開きます（進み具合を表示）。判定をすり抜けたときは［再スキャン］を押してください。2回目からはキャッシュ（`~/Library/Application Support/illustrator-scripts/FavoriteFontPicker-fonts.txt`）を使います
- 合成フォントと、環境にないフォントは一覧に出ません
- ［使用フォントをフォルダーに記録］は、一度保存したドキュメントで使えます
- ロック・非表示のテキストには適用しません

### 紹介記事

- [DTP Transit 別館｜note](https://note.com/dtp_tranist/n/ncf9ff6feebf0)

### 更新履歴

- v1.0.0 (20261001) : 初期バージョン（「よく使うフォントパネル」v1.1.2 をもとに改変）。2カラムのダイアログ、規格のチェックボックス（常に有効）、絞り込み欄、［PostScript名で表示］、［適用］ボタン、状態の保存、フォント数の変化による自動の読み直しを追加。モリサワの判定で「ud」を含む名前をすべて拾っていたのを前方一致に改めた。［ドキュメント］のフォントを規格違いの間引きで隠さないようにした
