# スウォッチの連続適用

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplySwatchesToSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/ApplySwatchesToSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

UIはありません。

選択中のオブジェクトに対して、スウォッチパネルで選択されているスウォッチを順に適用します。

- テキストが1つの場合には、文字単位で変更
- テキストが2つの場合には、テキストオブジェクトごとに変更

## 概要：

スウォッチパネルで選択されているスウォッチ、または全スウォッチの中からプロセスカラーを使い、
選択オブジェクトやテキストに順またはランダムでカラーを適用します。

## 処理の流れ：

1. ドキュメントと選択状態の確認
2. スウォッチの取得（未選択時はプロセスカラーからランダム取得）
3. テキスト1つ選択時は文字ごとにスウォッチを適用
4. 複数オブジェクト選択時は位置順に並べ替えてスウォッチを適用

## 対象：

TextFrame, PathItem, CompoundPathItem（内部のPathItem含む）

## 限定条件：

オブジェクトが選択されていること

## 謝辞：

sort_by_position.jsx（shspage氏）を参考にしました。
https://gist.github.com/shspage/02c6d8654cf6b3798b6c0b69d976a891

## 更新履歴

- v1.0.0（2024-11-03）初期バージョン
- v1.0.1（2025-06-25）スウォッチを選択していないときには、全スウォッチを対象に



