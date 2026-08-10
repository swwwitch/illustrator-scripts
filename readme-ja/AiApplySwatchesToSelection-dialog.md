# AiApplySwatchesToSelection-dialog

[![Direct](https://img.shields.io/badge/Direct%20Link-AiApplySwatchesToSelection-dialog.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/AiApplySwatchesToSelection-dialog.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiApplySwatchesToSelection-dialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトやテキストに、スウォッチや定義済みカラーを適用するモーダルダイアログです。
- 適用単位（オブジェクト／1文字／単語／行／段落）と適用順（そのまま／逆順／ランダム／完全ランダム）をラジオボタンで選択します。
- 「ランダム」はカラーの並びをシャッフルして繰り返し適用、「完全ランダム」は適用先ごとに毎回抽選するため繰り返しがありません。
- ラジオを変えるたびにライブプレビュー。［OK］で確定、［キャンセル］（Esc）で元に戻します。
- 開いた時点の選択スウォッチを■で取り込み、それ（スウォッチ名で参照）を適用に使用。1色でも選択があれば優先します。
- スウォッチ未選択時は自動カラー（CMYK は CM/CY/MY 生成、RGB は既定色）を使用。
- 「単語」は各行の先頭色が互い違いになるよう配色。
- モーダルダイアログはメインエンジンで動作するため、DOM 操作を直接実行します（BridgeTalk 委譲なし）。

### 紹介記事（note）

https://note.com/dtp_tranist/n/n5602f3084d2b

### スクリプト情報

- バージョン: v1.8.0
