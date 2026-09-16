# ドキュメントで使用しているフォント情報を書き出す

[![Direct Link](https://img.shields.io/badge/Direct%20Link-ExportFontInfoFromXMP.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ExportFontInfoFromXMP.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportFontInfoFromXMP.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

<img alt="" src="https://www.dtp-transit.jp/images/ss-476-456-72-20250713-081802.png" width="50%" />

## 概要

アクティブなIllustratorドキュメントに埋め込まれているXMPメタデータからすべてのフォント情報を抽出し、テキストファイル（タブ区切り）／CSV／Markdownで書き出すスクリプトです。

合成フォントの構成フォントも取り出せます。

未保存のドキュメントや、保存後に編集したドキュメントでも実行できます。この場合はXMPに無いフォントをドキュメント内のテキストから拾って補います。

## 主な機能

- TXT／CSV／Markdownの3形式に対応（3種類の一括書き出しも可能）
- 書き出し先を、デスクトップ／ドキュメントと同じ階層から選択
- 書き出し後に、書き出し先フォルダーを開く（初期値はON）
- 未保存・編集中のドキュメントでも実行可能（不足分をテキストから補完）
- 見つからないフォントだけに絞り込んで書き出し
- 合成フォントは、構成フォントの一覧も出力
- CSVはUTF-16（BOM付き）で出力
- Markdownはアンダースコア（`_`）のみエスケープ
- 同名ファイルが存在する場合は、連番を付けて自動的にリネーム
- 上下キーでラジオボタンを選択
- 日本語／英語UI

## 使い方

1. フォント情報を書き出したいドキュメントを開きます。
2. `ExportFontInfoFromXMP.jsx` を実行します。
3. ダイアログで書き出し形式・書き出し先・オプションを選択します。
4. ［OK］をクリックして実行します。

ファイル名は「ドキュメント名 + `_fontInfo`」に、選択した形式の拡張子を付けたものになります（例: `sample_fontInfo.csv`）。

## ［書き出し形式］パネル

| 項目 | 内容 |
| --- | --- |
| テキストファイル（.txt） | タブ区切りのテキストファイルを書き出します。 |
| CSVファイル（.csv） | UTF-16（BOM付き）のCSVファイルを書き出します。Excelでそのまま開けます。 |
| Markdownファイル（.md） | 見出し付きのMarkdownファイルを書き出します。 |
| すべて（3種類書き出し） | 上記3形式をまとめて書き出します。 |

## ［書き出し先］パネル

| 項目 | 内容 |
| --- | --- |
| デスクトップ | デスクトップに保存します。 |
| ファイルと同じ階層 | ドキュメントと同じフォルダーに保存します。未保存のドキュメントでは選べません。 |
| 書き出し後にフォルダーを開く | 書き出したあと、書き出し先フォルダーをFinder／エクスプローラーで開きます。初期値はONです。OFFのときは、書き出したファイル名をアラートで表示します。 |

## ［オプション］パネル

| 項目 | 内容 |
| --- | --- |
| 見つからないフォントのみ | インストールされていないフォントだけを書き出します。該当がない場合は、アラートを表示して書き出しません。初期値はOFFです。 |

## 出力例

Markdown

```markdown
### sw-B

- fontName: ATC-73772d42
- fontFace: 
- fontType: 合成フォント
- fileName: sw-B

#### 構成フォント

- RyoGothicStd-Bold.otf
- NotoSansCJKjp-Light.otf
- ヒラギノ角ゴシック W2.ttc
- RyoGothicStd-Heavy.otf
```

テキストファイル

```text
fontName:	ATC-73772d42
fontFamily:	sw-B
fontFace:	
fontType:	合成フォント
fileName:	sw-B
構成フォント：
・RyoGothicStd-Bold.otf
・NotoSansCJKjp-Light.otf
・ヒラギノ角ゴシック W2.ttc
・RyoGothicStd-Heavy.otf
```

## 設定変数

スクリプト冒頭の「ユーザー設定」で初期値を変更できます。

| 変数 | 初期値 | 内容 |
| --- | --- | --- |
| `FILENAME_SUFFIX` | `"_fontInfo"` | 出力ファイル名に付けるサフィックス |
| `SECTION_DIVIDER` | `"-----------------------------"` | テキストファイルの区切り線 |
| `OPEN_FOLDER_DEFAULT` | `true` | ［書き出し後にフォルダーを開く］の初期値 |
| `MISSING_ONLY_DEFAULT` | `false` | ［見つからないフォントのみ］の初期値 |

## 注意事項

- フォント情報は保存済みのXMPから取得します。未保存・編集中のドキュメントではXMPが前回保存時のままなので、不足分をドキュメント内のテキストから補います。
- テキストから補った分は、`version` と `fileName` が空になります。また合成フォントとしては扱われないため、構成フォントの一覧は出力されません。**これらの情報が必要な場合は、保存してから実行してください。**
- テキストの走査は文字単位で行うため、未保存・編集中の大きなドキュメントでは時間がかかります。保存済みのドキュメントでは走査しません。
- ドキュメントが開かれていない場合、フォント情報が見つからない場合は、アラートを表示して終了します。
- 合成フォントには `version` を出力しません。
- ［見つからないフォントのみ］の判定は、インストール済みフォントとの名前照合です。合成フォントは構成フォントの側で判定し、構成フォントが取得できない場合は判定できないため対象外になります。

## 紹介記事

[【Illustrator】ドキュメントで使用されているフォント情報を書き出す｜DTP Transit 別館](https://note.com/dtp_tranist/n/n16e7e95652b6)

## 更新履歴

- v1.0.3 (2026-09-17): 未保存・編集中のドキュメントに対応（XMPに無い分をテキストから補完）。［見つからないフォントのみ］を追加。ボタンエリアを左右中央に変更
- v1.0.2 (2026-08-06): ［書き出し後にフォルダーを開く］を追加。合成フォントの構成フォントが1件欠ける不具合、上下キーでの選択移動、XMLエンティティの復号、CSVのエスケープを修正
- v1.0.1 (2026-06-17): 書き出し先（デスクトップ／同じ階層）の選択、パネルレイアウト、未保存チェックを追加
- v1.0.0 (2025-05-10): 初期バージョン
