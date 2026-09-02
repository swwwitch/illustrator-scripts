# カテゴリ別ウェイト順にフォントを一覧表示し、フォント見本を一瞬で作成する

[![Direct](https://img.shields.io/badge/Direct%20Link-TypefaceSampler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/TypefaceSampler.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypefaceSampler.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

Illustrator で使用可能なフォントを、ファミリー（`font.family`）単位にまとめ、ウェイト（太さ）順・スタイル（装飾）順に並べてアートボード上に整列描画します。キーワード・ウェイト・種類で対象を絞り込めるので、必要な範囲だけのフォント見本をすぐに作れます。

<img alt="" src="https://www.dtp-transit.jp/images/ss-776-1010-72-20250713-082053.png" width="70%" />

### 主な機能

- ファミリー単位のグループ化と、ウェイト・スタイル評価によるグループ内の並べ替え
- キーワード検索（AND / OR / NOT / 先頭一致）による絞り込み
- ウェイト（5段階）と種類（5分類）のチェックボックスによる絞り込み
- 出力内容の切り替え（フォント名＋ウェイト／スタイル、PostScript名、アルファベット見本、数字見本、カスタムテキスト）
- 日本語／英語インターフェイス対応

### 使い方

1. ドキュメントを開いた状態でスクリプトを実行します。
2. ダイアログでキーワード・出力内容・絞り込み条件を指定します。
3. OK を押すと、アクティブなアートボードの左上から順に描画されます。

キーワードを空欄のままにすると全フォントが対象になるため、実行前に確認ダイアログが表示されます。フォント数によっては非常に時間がかかります。

### オプション

**キーワード**

検索対象は `font.name` / `font.family` / `font.style` の3つです。

| 記法 | 意味 | 入力例 |
| --- | --- | --- |
| スペース、`,` | OR（いずれかを含む） | `新ゴ 游` |
| `+` | AND（すべてを含む） | `新ゴ+游` |
| `^` | 先頭一致 | `^DIN` |
| `-` | 除外（NOT） | `Helvetica -Now` |

全角スペース・全角カンマが混ざっていても正しく処理されます。組み合わせも可能です（例：`^DIN+Bold -Condensed` → DIN で始まり Bold を含み、Condensed を含まない）。

**出力内容**

| 項目 | 描画される内容 |
| --- | --- |
| フォント名＋ウェイト／スタイル | `font.family` と `font.style` |
| PostScript名 | `font.name` |
| The quick brown fox… | アルファベットの見本 |
| 1234567890 | 数字の見本 |
| カスタム | 入力欄の文字列 |

**表示オプション**

| 項目 | 内容 |
| --- | --- |
| ウェイト数 | カテゴリー名の後ろにフォント数を併記 |
| ウェイト一覧 | オフにすると、カテゴリー名だけを1列に、そのカテゴリーで最も細いウェイトで組みます |
| 列数 | カテゴリーを並べる列数（↑↓キーで増減、shift で10単位） |
| スコア（検証用） | 並べ替えに使った評価値を併記（「フォント名＋ウェイト／スタイル」表示時のみ） |

**ウェイト**

`font.style` から5段階に分類して絞り込みます。`W3` / `W600` のような数値スタイル、`25 Ultra Light` のような先頭数値のスタイルにも対応します。

| 分類 | 該当するスタイルの例 |
| --- | --- |
| 超極細・極細 | Hairline、Ultra Thin、Thin |
| 細め | Ultra Light、Extra Light、Light |
| 標準 | Book、Normal、Regular、Roman |
| 中太 | Medium、SemiBold、DemiBold |
| 太字・極太 | Bold、Extra Bold、Heavy、Black、Ultra |

**種類**

`font.style` に含まれる装飾語で分類します。1つのフォントが複数の分類に該当することがあります。

| 分類 | 判定に使う語 |
| --- | --- |
| 基本 | Text、Headline |
| 狭める系 | Cond、Condensed、Compressed、Comp |
| 広げる系 | Expanded、Extended |
| 装飾・特殊用途 | Compact、Display |
| サイズ・プロポーション系 | Micro、Low、Wide |

ウェイト・種類とも、1つも選択していなければその条件は無視されます。両方を選択した場合は AND（両方を満たすフォントだけ）になります。

### 注意点

- キーワードを空欄にして実行すると、環境によっては数千書体が対象になり、処理に時間がかかります。
- 「基本」は `font.style` に Text / Headline を含むフォントのみが対象です。一般的な Regular や Bold は該当しません。
- 適用に失敗したフォントはスキップされ、ExtendScript のコンソールにログが出力されます。

### 紹介記事

- [【Illustrator】カテゴリ別ウエイト順にフォントを一覧表示し、フォント見本を一瞬で作成するスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n103ac6622657)

### 更新履歴

- v1.0.0 (20250420) : 初期バージョン
- v1.3.0 (20250508) : 評価ロジック修正、条件絞り込み機能強化
- v1.3.1 (20250706) : ローカライズ調整
- v1.3.2 (20260902) : ウェイト・種類による絞り込みを追加（TypefaceSampler-text.jsx を統合）
