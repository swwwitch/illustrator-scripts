# シンボルに登録して一致するオブジェクトを置き換え

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolizeAndReplace.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolizeAndReplace.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeAndReplace.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトをシンボルとして登録し、ドキュメント内の一致するオブジェクトをそのインスタンスに置き換え

### 主な機能

- シンボル名と 3×3 の基準点をダイアログで指定（既存シンボルと同じ名前は指定できません）
- テキストを選択した場合は、その文字列からシンボル名の初期値を作り、フォント・スタイル・文字列が一致するテキストフレームを対象にします（［フォントサイズ違いも対象にする］を ON にすると、サイズの違いを無視します）
- それ以外は、SmartEdit の一括選択で類似したオブジェクトを対象にします
- 複数のグループを選択した場合は、選択したすべてのグループを 1 つのシンボルに置き換えます（グループ以外を含む選択では実行できません）
- 指定した基準点で元オブジェクトの位置に揃えて、シンボルインスタンスに置き換えます
- 置換後は、新しいシンボルインスタンスを選択した状態のままにします
- ロック／非表示などで置換できなかった件数を報告します

### 紹介記事（note)

https://note.com/dtp_tranist/n/n650a4b91329d

### スクリプト情報

- バージョン: v1.0.1
