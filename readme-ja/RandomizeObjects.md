# RandomizeObjects

[![Direct](https://img.shields.io/badge/Direct%20Link-RandomizeObjects.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RandomizeObjects.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RandomizeObjects.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 更新日

- 2026-03-05（v2.2：［重なりを避ける］の配置ロジックを汎用関数として抽出）
- 2026-02-27（v2.1：完全シャッフル＝ランダム色生成（CMYKのKは0–30）／カラーは［ランダム］で実行）

### 概要

- 選択したオブジェクトをランダムに移動・変形・回転・不透明度を変更するスクリプト
- UIから各種パラメータを指定し、即時プレビューで結果を確認可能
- カラーの通常シャッフル／完全シャッフルに対応
- [リセット]でダイアログ起動前の状態に復元
- ［重なりを避ける］の配置ロジックを他スクリプトへ流用しやすい形で関数化

### 主な機能

- 移動距離を横・縦方向に個別設定、または連動
- 中央揃えオプションでオブジェクトを一箇所に集約
- 拡大縮小（変形）、回転、不透明度のランダム化
- プレビュー反映、キャンセルで元の状態に完全リセット
- UI状態の同期安定性向上（チェックON/OFFと入力欄enabledの整合性を統一）
- プレビューを毎回ベース状態から再計算（積み増し挙動を解消）

### 処理の流れ

1. ドキュメントと選択状態を確認
2. ダイアログを表示（移動距離・変形率・回転・不透明度を入力）
3. 入力値をもとにランダム変形や移動を即時プレビュー
4. OKで確定、キャンセルでダイアログ起動前にリセット

### スクリプト情報

- バージョン: v2.2
