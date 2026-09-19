# 開いているすべてのドキュメントへ同じ位置にペースト

[![Direct](https://img.shields.io/badge/Direct%20Link-PasteInAllDocs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/PasteInAllDocs.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PasteInAllDocs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

コピー済みのオブジェクトを、開いているすべてのドキュメントへ同じ位置に貼り付けます。

「同じ位置にペースト」（pasteInPlace）を使うため、各ドキュメントの座標にあわせて配置されます。

### 使い方

1. 貼り付けたいオブジェクトをコピーします。
2. 貼り付け先のドキュメントをすべて開いた状態で、スクリプトを実行します。

### 注意点

- 事前にコピーしていない場合は、警告を表示して終了します。
- コピー元のドキュメントには複製されません。
- 貼り付け先では、コピー元と同じ番号のアートボードをアクティブにしてからペーストします。
- 貼り付け先のアクティブレイヤーがロックまたは非表示の場合は一時的に解除し、ペースト後に元へ戻します。
- ペーストできなかったドキュメントがあれば、最後にまとめて名前を表示します。
- 処理中、各ドキュメントの選択は解除されます。
- 環境の言語に応じて、メッセージを日本語／英語で表示します。

### 紹介記事

[【Illustrator】すべてのドキュメントにペースト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n04535658c7f6)

### 更新履歴

- v1.1.0 (2026-09-19): メッセージを日本語／英語に対応。コピー元への複製、アートボード違いによる位置ずれ、ペースト失敗を検知できない問題を修正。紹介記事へのリンクを追加
- v1.0
