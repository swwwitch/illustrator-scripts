# 複数ドキュメントの選択オブジェクトを同じ座標に揃える

[![Direct](https://img.shields.io/badge/Direct%20Link-SyncSelectionPosition.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/SyncSelectionPosition.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SyncSelectionPosition.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

最前面のドキュメントの選択範囲の左上を基準に、ほかの開いているドキュメントの選択オブジェクトを同じ座標へ移動します。

### 使い方

1. 2つ以上のドキュメントを開きます。
2. 基準にしたいドキュメントを最前面にして、そこでオブジェクトを選択します。
3. 位置を合わせたいドキュメントでもオブジェクトを選択しておきます。
4. スクリプトを実行します。

### 注意点

- 開いているドキュメントが2つ未満のときは、警告を表示して終了します。
- 最前面のドキュメントで何も選択されていないときも、警告を表示して終了します。
- 基準にするのは選択範囲の左上（Left / Top）の座標です。
- 選択のないドキュメントは、何もせずスキップします。
- アラートの文言は、Illustratorの言語環境に応じて日本語／英語を切り替えます。

### 紹介記事

- [【Illustrator】複数ドキュメントの選択オブジェクトを同じ座標に揃えるスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n1f8155daeac4)

### 更新履歴

- v1.0 (20251227) : 初期バージョン
- v1.0.1 (20260903) : アラートの文言をLABELS＋`getLabel()`に整理して日英切り替えに対応
