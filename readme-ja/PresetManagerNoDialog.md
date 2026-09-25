# 決めておいた環境設定一式をダイアログなしで適用

[![Direct](https://img.shields.io/badge/Direct%20Link-PresetManagerNoDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManagerNoDialog.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerNoDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

PresetManager の［プリセット1］と同じ環境設定一式を、ダイアログを表示せずにまとめて適用します。

### 使い方

スクリプトを実行します。

### 注意点

- 適用する内容はスクリプト内の `PREFERENCES` に、環境設定キーと値を1行ずつ書いてあります。変更したい場合はスクリプトを編集してください。
- バージョンの違いにより、一部のキーは無視される場合があります。存在しないキーは安全に読み飛ばします。

### 更新履歴

- v1.1.0 (2026-09-25) ［プリセット1］だけに絞り、`minimal` / `full` と `ACTIVE_PRESET` の切り替えを廃止。環境設定キーと値を1つの表にまとめて簡素化
- v1.0.1 (2026-09-25) `preset1` を PresetManager の［プリセット1］と完全に一致させた。「カンバス上でロック解除」を追加し、スマートガイド／グリッドのスナップ設定を外した
- v1.0 (2026-09-19) PresetManagerNoDialogFull / PresetManagerPreset1 を統合し、`ACTIVE_PRESET` での切り替え方式に変更
- v1.0
