# 重なったオブジェクトを横または縦へ並べ直す

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAlignAndTile.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignAndTile.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 重なって配置されたオブジェクトを、横方向または縦方向へ指定した間隔で並べ直します。行数・列数を指定すればタイル状にも配置できます。
- Illustratorで設定したキーオブジェクトを自動判定し、その位置を動かさないまま並べ替えられます。
- ダイアログの操作はそのままプレビューに反映され、Undo履歴を汚さずに確認できます。
- 更新日：2026-09-05
- `SmartAlignAndTile-yoko`（横）と `SmartAlignAndTile-tate`（縦）を1本に統合したスクリプトです。

### 主な機能

- 並べる方向を横・縦から選択
- 行数（横）／列数（縦）を指定して複数行・複数列に分割
- 横・縦の間隔を個別に指定（「連動」で横の値を縦にも適用）
- 定規の単位に追従（mm、pt、px、Q/H など）
- グリッド配置（最大サイズをセルとして等幅・等高に配置）
- 天地の揃え（上・中央・下・なし）と左右の揃え（左・中央・右・なし）
- キーオブジェクトを基準にした配置（自動判定、未検出のときはディム）
- ランダム配置（全体の左上位置は保持）
- 数値欄を↑↓キーで増減（Shift+↑↓で10単位にスナップ）
- Undoを汚さないプレビューと、OK後の1回のUndoでの取り消し

### 使い方

1. 並べ替えたいオブジェクトを選択します。基準にしたいオブジェクトがある場合は、選択したうえでそのオブジェクトをもう一度クリックしてキーオブジェクトに設定します。
2. スクリプトを実行します。
3. 方向・行数（列数）・間隔・揃え・オプションを設定します。操作するたびにプレビューが更新されます。
4. ［OK］で確定します。取り消すときはUndo（⌘Z）1回で戻せます。

### オプション

- **方向**：横＝左から右へ並べ、行数で折り返します。縦＝上から下へ並べ、列数で折り返します。
- **揃え（上下）**：上・中央・下のいずれかに揃えます。「なし」は縦位置を変えず、行・列の送りだけを適用します。
- **揃え（左右）**：左・中央・右のいずれかに揃えます。「なし」は横位置を変えません。
- **キーオブジェクトを基準**：判定できたキーオブジェクトの位置を保ったまま、ほかのオブジェクトを並べます。判定できないときはディムされます。
- **プレビュー境界を使用**：ONで線幅や効果を含む境界（visibleBounds）、OFFで幾何境界（geometricBounds）を基準にします。
- **グリッド**：もっとも大きいオブジェクトのサイズをセルとして、等幅・等高のグリッドに配置します。ONにすると天地・左右とも中央揃えが既定になります。
- **ランダム**：並び順をランダムにします。全体の左上位置は変わりません。
- **連動**：横の間隔と同じ値を縦にも適用します（ONのあいだ縦の入力欄はディム）。

### 注意点

- 並べる方向の揃え（横なら左右、縦なら上下）は、セルがオブジェクトより大きくなるグリッド時にだけ意味を持つため、グリッドOFFではディムされます。
- キーオブジェクトの判定は、4方向の整列コマンドを試して「どの向きでも動かなかったオブジェクト」を特定する方式です。判定のあいだオブジェクトが一時的に動き、Undo履歴にその痕跡が残ります。
- 選択が1つだけのとき、または候補が複数あるときはキーオブジェクトを判定できません。
- 行数・列数で割り切れない場合、最後の行（列）のオブジェクト数は少なくなります。

### オリジナルアイデア

John Wundes
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

### 紹介記事

- [【Illustrator】複数のオブジェクトを整列・タイル配置するスクリプト updated｜DTP Transit 別館](https://note.com/dtp_tranist/n/nf426908d8bcd)

### 更新履歴

- v2.0 (20260905) : `SmartAlignAndTile-yoko`（v1.7.1）と `SmartAlignAndTile-tate`（v1.8）を統合し、方向（横／縦）の切り替えに対応。揃えに「なし」を追加し、行・列の基準を最大サイズに統一

---

### スクリプト情報

- バージョン: v2.0
