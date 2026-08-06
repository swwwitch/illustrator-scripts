# アートボード名の一括変更

[![Direct](https://img.shields.io/badge/Direct%20Link-RenameArtboardsPlus.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RenameArtboardsPlus.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

![](https://www.dtp-transit.jp/images/ss-1210-1120-72-20250713-080350.png)

- 接頭辞・接尾辞・元のアートボード名の有無を指定可能
- アートボード名と番号の組み合わせ（番号、名称、番号-名称、番号_名称）に対応
- 連番は「数字」「アルファベット（大文字／小文字）」に対応
- 「開始番号」の入力文字列から桁数（ゼロ埋め）を自動判定
- 「ファイル名の参照」や「区切り文字」など柔軟な命名に対応
- プレビュー機能により変更内容を事前確認可能（最大15件）
- OK／適用／キャンセルボタンを左右に美しく配置
- プリセット書き出し・組み込みプリセット選択に対応（*.txt形式、labelは単一文字列）

## 更新履歴

- v1.3.1（2026-08-06）プリセット書き出しをUTF-8で保存するよう修正（日本語ラベルの文字化けを防止）。連番形式が「なし」のとき、効果のない接尾辞の区切り文字をディムするよう変更。開始番号・増分に数字以外が混じる入力を弾き、前後の空白を無視するよう修正。リネーム中にエラーが起きたときはダイアログを閉じないよう変更。パネルの余白・間隔を共通設定（setupPanel / PANEL_MARGINS）に統一。変数・関数・パネル名を実状に合わせて整理し、LABELSをカテゴリ別に構造化、全関数にJSDocを追加
- v1.3.0（2026-05-09）名前が空になる場合の警告を追加。区切り文字のラジオボタンに値を関連付け、配列順序への依存を解消。プリセット書き出し時に文字列をエスケープするよう修正。開始番号の初期値を「001」に変更
- v1.2.0（2026-05-07）ローカライズを刷新（L / labelText / labelWithCount）し、内部キー化により英語ロケールに対応。連番形式に「なし」を追加。［適用］ボタンを削除。main() をUIビルダー／プリセットI/O／純粋関数に分割
- v1.1（2025-04-30）開始番号から桁数（ゼロ埋め）を自動判定する処理を追加。プリセットのラベルを簡素化し、ES3対応を強化
- v1.0（2025-04-20）初期バージョン

### note

- [【Illustrator】連番や指定文字を使ってアートボード名前を一括変更｜DTP Transit 別館](https://note.com/dtp_tranist/n/n80f9534bc6fb)