# 字形の境界を切り替えながら垂直方向に整列

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartVerticalAlign.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartVerticalAlign.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartVerticalAlign.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ポイント文字およびエリア内文字の［字形の境界に整列］を切り替えながら、垂直方向（上・中央・下）に整列します。チェックボックスやラジオボタンの操作はそのつどプレビューへ反映されるため、字形の境界を含めた場合と含めない場合を見比べながら決められます。

### 主な機能

- ポイント文字／エリア内文字それぞれの［字形の境界に整列］をON/OFF
- 整列位置（上・中央・下）を選択
- ［プレビュー境界］のON/OFFで、線幅・効果を境界に含めるかを切替
- チェックボックス・ラジオボタンの操作はすべてプレビューに即時反映
- ホットキー（T／M／B）で整列位置を切替
- 日本語／英語UI

### 使い方

1. 整列したいオブジェクトを選択します。
2. スクリプトを実行します。
3. ［字形の境界に整列］と整列位置を指定し、プレビューを確認して［OK］をクリックします。

### オプション

**字形の境界に整列**

| 項目 | 内容 |
| --- | --- |
| ポイント文字 | ポイント文字の［字形の境界に整列］（環境設定 `EnableActualPointTextSpaceAlign`） |
| エリア内文字 | エリア内文字の［字形の境界に整列］（環境設定 `EnableActualAreaTextSpaceAlign`） |

選択がポイント文字のときはエリア内文字が、エリア内文字のときはポイント文字がディム表示になります。

**整列**

| 項目 | 内容 |
| --- | --- |
| 上 | 上揃え |
| 中央 | 中央揃え |
| 下 | 下揃え |

**プレビュー境界**

| 項目 | 内容 |
| --- | --- |
| プレビュー境界 | ONで線幅・効果を境界に含めます（環境設定 `includeStrokeInBounds`）。デフォルトOFF |

**ホットキー**

| キー | 内容 |
| --- | --- |
| T / M / B | 上／中央／下 |

### 注意点

- 整列対象はテキスト・パス・グループ・複合パスです。
- 選択内容によって初期値が変わります。テキストのみの場合は「下」、テキストと図形が混在する場合や図形のみの場合は「中央」が選ばれます。
- チェックボックスの操作はIllustratorの環境設定を直接書き換えます。［キャンセル］で元の状態には戻りません。
- 整列はIllustratorのメニューコマンドで実行するため、キー入力を含むすべての操作が取り消し履歴に残ります。

### 紹介記事

https://note.com/dtp_tranist/n/n9ee716675032

### 更新履歴

- v1.1.1 (2026-09-03) ［字形の境界に整列］をOFFに戻したときもプレビューに反映されるよう修正、整列位置・ショートカット・環境設定キーをテーブル化し、命名と構成を整理
- v1.1 (2025-08-04) ダイアログボックスを開くときのロジックを調整
- 初版 (2025-08-04)

### スクリプト情報

- バージョン: v1.1.1
- 初回リリース: 2025-08-04
- 最終更新: 2026-09-03
