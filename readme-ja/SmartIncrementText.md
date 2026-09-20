# テキスト内の数字・英字を増分しながら複製

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartIncrementText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartIncrementText.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartIncrementText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストフレーム内の数字・英字・日付・時刻を検出し、値を増分しながら下方向へ複製します。ダイアログでの操作はそのままプレビューに反映されます。

### 主な機能

- 数字（`01`、`2025` など）と英字1文字（`A`〜`Z` / `a`〜`z`）の増分
- 日付（`2026年9月20日` / `2026/09/20` / `2026.9.20`）は年・月・日のどれを増分するかを選べ、暦として正しく繰り上がります。括弧内の曜日も追従します
- 時刻（`19:00`）は時・分のどちらを増分するかを選べ、24時間で繰り上がります
- 増分対象が複数あるときは、ラジオボタンで選択（`数字1` / `数字2` / `英字1` など）
- 開始値の指定、ゼロ埋め（最終値が桁上がりする場合は桁数を自動で広げます）
- ［OK］で複製を1つのテキストに結合（行送り＝文字サイズ＋［アキ］）
- ↑↓キーで数値を増減（Shiftで10、Optionで0.1）
- 前回のダイアログ位置を記憶

### 使い方

1. 元になるテキストフレームを選択します。
2. スクリプトを実行します。
3. ［複製数］［増分］［アキ］などを指定します（プレビューで確認できます）。
4. ［OK］で確定します。

### 主な設定

| 項目 | 内容 |
| --- | --- |
| 複製数 | 作る複製の数（元のテキストは含みません） |
| 増分 | 1つ進むごとに足す数（負の値で減らせます） |
| アキ | 複製どうしのアキ。文字サイズに加算されます |
| 増分対象 | テキストの中で増やす箇所 |
| 開始値 | 元のテキストの値ではなく、指定した値から始めます |
| ゼロ埋め | 元の桁数に合わせて頭に0を足します（最終値で桁が増えるときは広げます） |
| 1つのテキストに結合 | ［OK］のときに、複製を改行でつないで1つのテキストにまとめます |

### 注意点

- 英字の増分は1文字（`A`〜`Z`）のみに対応します。`A1` は対象になりますが、`AB1` や `Ver1` の `AB` / `Ver` は対象になりません。
- すでにある連番を振り直したい場合は SmartRenumber.jsx を使用してください。

### 紹介記事

- [【Illustrator】連番を一括生成するスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n5f25ed17b123)

### 更新履歴

- v2.0.0
- v2.0.2 (20260920) : 項目名を［アキ］［開始値］［1つのテキストに結合］に変更、ゼロ埋めで負の値が崩れる問題を修正、増分対象が1つのときは行を表示しない、開始値が不正なときは［OK］を無効化
