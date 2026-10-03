# グループ・複合パス・複合シェイプ・クリッピングマスクを入れ子ごと解除

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseGroupsAndMasks.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/ReleaseGroupsAndMasks.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseGroupsAndMasks.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトのグループ・複合パス・複合シェイプ・クリッピングマスクを、入れ子のものまでまとめて解除します。
- クリップグループは、マスクパスと内容の両方を残して解除し、マスクパスに K100・不透明度15% の塗りを設定します（ダイアログは出しません）。
- グループを含むグループがあるときだけ、ダイアログで［すべて解除］か［1階層だけ解除］を選べます。

### 主な機能

- グループ・複合パス・複合シェイプを、なくなるまで繰り返し解除
- クリップグループは単純に解除し、残したマスクパスに K100・不透明度15% の塗りを設定
- 入れ子のグループは［すべて解除］か［1階層だけ解除］を選択
- 種類の違うオブジェクトを混ぜて選んでも、1つずつ選び直して解除
- 日本語／英語 UI

### 使い方

1. 解除したいオブジェクトを選択する
2. スクリプトを実行する
3. ダイアログが出たら、解除する深さを選んで［OK］

### オプション

| パネル | 項目 | 内容 |
| --- | --- | --- |
| 入れ子のグループ | すべて解除 | 入れ子のグループ・複合パス・複合シェイプを残らず解除する（初期値） |
| | 1階層だけ解除 | 選択したものを1回だけ解除し、中のものは残す |

ダイアログは、グループを含むグループがあるときだけ表示します。

クリップグループの解除方法と塗りの有無は、スクリプト冒頭の `USER_DEFAULTS` で変えられます。

| 設定 | 値 |
| --- | --- |
| `clipReleaseMode` | `"simple"`（マスクパスと内容を残す、初期値）／ `"removePath"`（マスクパスを削除）／ `"removeContent"`（マスク内容を削除） |
| `applyMaskFill` | `true`（初期値）で、残したマスクパスに塗りを設定 |

### 注意点

- ［すべて解除］では、マスク内容の中にあるグループや複合パスも解除します。残したマスクパスが複合パスのときは、複合パスのまま残します。
- 複合シェイプは［複合シェイプを解除］のアクションで解除します。ブレンドやエンベロープは解除しません。
- 入れ子のクリップグループは、マスクとして扱わずにグループとして解除します（マスクパスは塗り・線のない状態で残ります）。
- 処理のあとは、解除で出てきたオブジェクトを選択した状態になります。

### 更新履歴

- v1.0.0 (20261004) : 初期バージョン
