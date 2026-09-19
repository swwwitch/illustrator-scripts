# 決めておいた環境設定一式をダイアログなしで適用

[![Direct](https://img.shields.io/badge/Direct%20Link-PresetManagerNoDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManagerNoDialog.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerNoDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

あらかじめ決めた環境設定一式を、ダイアログを表示せずにまとめて適用します。
冒頭の `ACTIVE_PRESET` で、適用するプリセットを切り替えます。

### プリセット

| 値 | 内容 |
| --- | --- |
| `minimal` | 定番だけを手早く。19項目 |
| `full` | 環境設定をひととおり押さえた版。31項目。「ブラックのアピアランス」や定規の単位など、再起動が必要な項目を含む |
| `preset1` | PresetManager の［プリセット1］と同じ内容。38項目。ガイド・スマートガイド・アートボードのハイライトを含む |

### 使い方

1. スクリプトの `ACTIVE_PRESET` に `"minimal"` / `"full"` / `"preset1"` のいずれかを指定します。
2. スクリプトを実行します。

### 注意点

- 適用する内容はスクリプト内の `PRESET_STATES` に書かれています。変更したい場合はスクリプトを編集してください。
- プリセットに書かれていない項目には触れません。Illustrator の現在値がそのまま残ります。
- 環境設定キーは `PREFERENCE_BINDINGS` の対応表にまとめてあります。項目を増やすときはこの表に1行足します。
- バージョンの違いにより、一部のキーは無視される場合があります。存在しないキーは安全に読み飛ばします。
- 定規の単位、ユーザーインターフェイスの明るさ、［シェイプ形成ツール］の［次のカラー］は、反映に Illustrator の再起動が必要です。

### 更新履歴

- v1.0 (2026-09-19) PresetManagerNoDialogFull / PresetManagerPreset1 を統合し、`ACTIVE_PRESET` での切り替え方式に変更
- v1.0
