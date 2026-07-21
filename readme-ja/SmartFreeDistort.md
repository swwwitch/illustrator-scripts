# SmartFreeDistort.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-SmartFreeDistort.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartFreeDistort.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## 概要

選択オブジェクトに Illustrator の「自由変形」ライブ効果を適用するスクリプトです。台形・平行四辺形・三角形・対角線の全18プリセットをアイコンから選択できます。

調整可能なプリセットでは変形量と強度を設定し、Undoベースのプレビューで結果を確認してから適用できます。

## 主な機能

- すべてのプリセットを形状アイコンで表示
- 台形4種類
- 平行四辺形8種類（4基準点 × 左右／上下）
- 三角形4種類
- 対角線2種類
- 変形量を `0.00`〜`0.49` の範囲で調整
- 強度を「マイルド」「ノーマル」「ブースト」から選択
- Undoベースのプレビュー
- 複数オブジェクトへの一括適用
- 日本語／英語UI

## 使い方

1. 自由変形を適用するオブジェクトを選択します。
2. `SmartFreeDistort.jsx` を実行します。
3. ダイアログでプリセットを選択します。
4. 台形または平行四辺形では、変形量と強度を調整します。
5. 必要に応じて「プレビュー」を有効にし、結果を確認します。
6. 「OK」をクリックしてライブ効果を適用します。

## プリセット

| 種類 | プリセット数 | 変形量・強度 |
| --- | ---: | --- |
| 台形 | 4 | 使用する |
| 平行四辺形 | 8 | 使用する |
| 三角形 | 4 | 使用しない |
| 対角線 | 2 | 使用しない |

## 強度

| 設定 | 倍率 |
| --- | ---: |
| マイルド | 0.25 |
| ノーマル | 0.5 |
| ブースト | 1.0 |

選択範囲にテキストが含まれる場合は、文字の崩れを抑えるため「マイルド」が初期値になります。

## 注意事項

- Illustrator 2024〜2026に対応しています。
- 効果を適用できない選択項目は処理対象から除外されます。
- 複数選択時は、同じライブ効果を各オブジェクトへ個別に適用します。
- プレビューは Illustrator のUndo履歴を利用します。他の操作と混在すると、履歴が想定どおりにならない場合があります。
- 三角形と対角線では、変形量と強度の設定は使用されません。

## 紹介記事

[【Illustrator】自由変形を手軽に適用するスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n15a7ae196a23)

## 更新履歴

- v1.5.4 (2026-07-21): 日本語・英語READMEを追加し、スクリプトの基本情報からリンク
- v1.5.3 (2026-07-21): 概要を現在のプリセット数と適用フローに合わせて更新
- v1.5.2 (2026-07-21): プレビュー解除、対象判定、部分適用時の通知を改善
- v1.5.1 (2026-07-21): 全プリセットをアイコン化し、平行四辺形を8種類へ拡張
- v1.5.0 (2026-07-21): プレビュー、対象抽出、プリセット定義を整理
- v1.1.1 (2026-04-24): 初期バージョン
