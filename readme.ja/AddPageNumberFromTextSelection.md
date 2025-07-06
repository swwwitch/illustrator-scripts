# ページ番号を挿入

[![Direct](https://img.shields.io/badge/Direct%20Link-AddPageNumberFromTextSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/AddPageNumberFromTextSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- _pagenumberレイヤーで選択したテキストを基準に、すべてのアートボードにページ番号テキストを複製・配置するIllustrator用スクリプトです。
- 開始番号、接頭辞、接尾辞、ゼロ埋め、総ページ数表示のカスタマイズが可能です。

![](../png/ss-672-346-72-20250629-205331.png)

### 主な機能

- 選択テキストを基にページ番号を生成
- 開始番号、接頭辞、接尾辞の指定
- ゼロパディング（ゼロ埋め）対応
- 総ページ数の表示オプション
- プレビュー機能
- 日本語／英語インターフェース対応

### 処理の流れ

1. _pagenumberレイヤーにテキストを選択
2. ダイアログで開始番号や接頭辞などを設定
3. プレビューを確認
4. OKでページ番号テキストを全アートボードに配置

### 更新履歴

- v1.0.0 (20240401) : 初期バージョン
- v1.0.1 (20240405) : テキスト複製ロジック修正
- v1.0.2 (20240410) : ゼロ埋め・接頭辞・総ページ数表示追加
- v1.0.3 (20240415) : プレビュー機能追加
- v1.0.4 (20240420) : 「001」形式のゼロ埋め対応
- v1.0.5 (20240425) : 接尾辞フィールド追加、UI改善

### 課題：

- プレビュー時、元のテキストが残ってしまい重複して見える問題があります。OKボタンを押すと消えます。



