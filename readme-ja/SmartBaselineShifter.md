# 特定の文字のベースラインシフトを調整

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartBaselineShifter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartBaselineShifter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要：

- 選択したテキストフレーム内の指定文字にベースラインシフトを個別適用（単位は環境設定の「単位」→「東アジア言語」）
- ダイアログで対象文字列とシフト量を設定し、即時プレビュー可能

<img alt="" src="https://www.dtp-transit.jp/images/ss-742-402-72-20250716-205309.png" width="70%" />

### 主な機能：

- 対象文字列の指定
- シフト量の整数・小数指定
- すべてをリセット
- 即時プレビュー、元に戻す操作
- 日本語／英語UI対応

#### 対象文字列の指定

- 選択しているテキストから、数字やスペースなどを除いたものが自動的に入ります。
- 編集可能です。
- 「:-」と入力すれば「:」と「-」の両方が対象になります。

#### 値の変更

- ↑↓キーで±1増減
- shiftキーを併用すると±10増減
- optionキーを併用すると±0.1増減

### 処理の流れ：

1. テキストフレームを選択
2. ダイアログで設定
3. プレビュー確認
4. OKで確定、キャンセルで元に戻す

### オリジナル、謝辞：

Egor Chistyakov https://x.com/tchegr

### note

https://note.com/dtp_tranist/n/n5e41727cf265

### 更新履歴：

- v1.0 (20240629) : 初期バージョン
- v1.3 (20240629) : +/-ボタン追加
- v1.4 (20240629) : ダイアログ2カラム化、正規表現対応
- v1.5 (20240630) : TextRange選択用関数追加
- v1.6 (20240630) : 正規表現対応削除、微調整
- v1.7 (20240630) : ↑↓キーでの値の変更機能追加、UIの再設計
- v1.8 (20250720): 自動計算機能を追加
- v2.2.2 (20260921): シフト量を環境設定の単位で扱うよう修正、シフト量に小数を入力できない不具合を修正、対象文字の初期値から空白・改行を除外、自動計算で対象文字が見つからないときに通知
