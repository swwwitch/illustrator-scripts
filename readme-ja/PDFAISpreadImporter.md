# PDF/AIの各ページを新規ドキュメントのアートボードに配置

[![Direct](https://img.shields.io/badge/Direct%20Link-PDFAISpreadImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/PDFAISpreadImporter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAISpreadImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、新規ドキュメント上に各ページを個別のアートボードとして配置します。

横長ページは見開きとして自動判定し、左右2つのアートボードに分割して配置します。

<img alt="PDF/AI見開き配置ダイアログの外観" src="../png/ss-870-628-144-20260925-061426.png" width="50%" />

### 主な機能

- ページ範囲の指定（全ページ／先頭ページのみ／指定ページ）
- 横長ページを見開きとして自動判定し、左右に分割
- 偶数ページの位置を右／左から選択
- PDFのトリミング設定（アート／トリミング／仕上がり／裁ち落とし）を選択
- 新規ドキュメントのカラーモードを CMYK / RGB から選択

### 使い方

1. スクリプトを実行します。
2. 読み込む PDF/AI ファイルを指定します（選択中の配置画像があればそれを使います）。
3. ページ範囲、偶数ページの位置、トリミング設定、カラーモードを指定します。
4. 実行すると、新規ドキュメントにアートボードが並びます。

### オプション

#### トリミング

PDF のどのボックスを基準に配置するかを選びます。AI ファイルでは使いません（PDF のときだけ選べます）。

<img alt="トリミングの選択肢" src="../png/ss-412-240-144-20260925-061158.png" width="25%" />

| 選択肢 | 使われるボックス | 内容 |
|---|---|---|
| アート | ArtBox | 作成者が指定したアートの範囲。指定がない PDF ではトリミングと同じ |
| トリミング | CropBox | 表示・印刷される範囲（Acrobat で見えている範囲） |
| 仕上がり（初期値） | TrimBox | 断裁後の仕上がりサイズ。指定がない PDF ではトリミングと同じ |
| 裁ち落とし | BleedBox | 仕上がりに塗り足しを加えた範囲。指定がない PDF ではトリミングと同じ |

入稿用の PDF（トンボや塗り足し付き）を見開きで分割するときは、左右が仕上がり位置で分かれる［仕上がり］を使います。

### 注意点

- 選択中の配置画像、または指定ファイルからページ数を推定し、ページ範囲に反映します。
- アートボード間隔は 100 pt 固定です。
- ラスタライズ効果解像度は 300 ppi 固定です。
- 各ページを分割せずに配置したい場合は PDFAIImporter.jsx を使用してください。

### 紹介記事

[【Illustrator】見開きのPDFを、片ページごと読み込みたい｜DTP Transit 別館](https://note.com/dtp_tranist/n/n5514d9f2c5f8)

### 更新履歴

- v1.1.2 (2026-09-25): 項目名にコロンを追加し、トリミングとカラーモードにラベル・ツールチップを追加。ボタン名を［ファイルを選択...］、パネル名を［ページ］に変更。トリミング指定の値の誤りを修正し、「仕上がり」「裁ち落とし」「アート」が選んだとおりのボックスで配置されるように。ページ数が多いときにキャンバスからはみ出してエラーになる問題を修正（右端で次の行へ折り返す）。内部の重複を整理
- v1.1.1 (2026-09-19): 紹介記事へのリンクを追加、各項目にツールチップを追加。内部の命名と構造を整理
- v1.1.0 (2026-03-18)
