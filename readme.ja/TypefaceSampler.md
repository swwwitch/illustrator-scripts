# カテゴリ別ウエイト順にフォントを一覧表示し、フォント見本を一瞬で作成する

[![Direct](https://img.shields.io/badge/Direct%20Link-TypefaceSampler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/TypefaceSampler.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Illustrator で使用可能なフォントを一覧表示し、ウエイト（太さ）順やスタイル（装飾）順に従ってアートボード上に整然と描画します。

### 機能概要：

- font.style をもとにウエイトや装飾キーワードを判定
- Ultra Light〜Black などのウエイト分類は weightGroups により定義
- Italic や Condensed、Wide などの装飾語には加点処理
- "25 Ultra Light" や "W300" のように数値スタイルにも対応（数値優先）
- 表示テキストにスコア（rank）を含める debug モードあり
- キーワード検索によるフォントの絞り込みに対応：
  ● 検索対象：font.name / font.family / font.style
  ● 入力例と意味：
    ・`新ゴ 游`           → OR（新ゴ または 游 を含む）
    ・`^DIN+Bold`         → AND（DINで始まり、Boldを含む）
    ・`Helvetica -Now`    → NOT（Helvetica を含み、Now を含まない）
    ・`新ゴ+游 -Light`     → 新ゴかつ游を含み、Lightを含まない
    ・`^DIN+Bold -Condensed` → 複合条件（先頭一致＋AND＋NOT）
  ● 構文仕様：
    - 「^」：先頭一致
    - 「+」：AND条件
    - 「,」またはスペース：OR条件（全角スペース・カンマもOK）
    - 「-」：除外（NOT条件）
  ● スペース・カンマ・全角空白混在でも正しく処理されます

### 処理の流れ：

1. ダイアログで出力形式・列数・カテゴリ分けなどのオプションを取得
2. Illustrator が現在利用可能なすべてのフォントから条件に合致するフォントを収集
3. 評価スコア順に並べてアートボード左上から描画

作成日：2025-04-20
最終更新日：2025-05-08 17:05

[【Illustrator】カテゴリ別ウエイト順にフォントを一覧表示し、フォント見本を一瞬で作成するスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n103ac6622657)