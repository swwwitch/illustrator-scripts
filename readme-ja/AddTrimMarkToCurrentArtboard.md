# AddTrimMarkToCurrentArtboard

[![Direct](https://img.shields.io/badge/Direct%20Link-AddTrimMarkToCurrentArtboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AddTrimMarkToCurrentArtboard.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMarkToCurrentArtboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 常に現在のアートボードに対して、日本式トンボを作成するIllustrator用スクリプトです。
- 専用の「トンボ」レイヤーを最初に取得または作成し、その上にアートボード矩形を1つだけ作成します。
- その矩形を元にトリムマークを作成し、同じオブジェクトをそのままガイド化します。
- 「トンボ」レイヤーがロックされている場合はいったん解除して処理し、終了時に元のロック状態へ戻します。

### 主な機能

- 常に現在のアートボードを対象にトリムマークを作成
- 「トンボ」レイヤーを自動取得／未存在時は新規作成
- アートボード矩形1つのみを使用（複製なし）
- 同一オブジェクトでトンボ生成とガイド化を完結
- 「トンボ」レイヤーの元のロック状態を維持

### 処理の流れ

1. 「トンボ」レイヤーを取得し、なければ新規作成
2. 環境設定で日本式トンボをONに設定
3. 「トンボ」レイヤー上にアートボード矩形を1つ作成
4. その矩形を元にトリムマーク作成メニューを実行
5. 同じ矩形オブジェクトをガイド化
6. finally で選択解除を行い、「トンボ」レイヤーのロック状態を元に戻す

### 更新履歴

- v1.0 (20250205) : 初期バージョン
- v1.1 (20260401) : 「トンボ」レイヤーがロックされている場合はいったん解除して処理し、終了時に元のロック状態へ戻すように変更

### スクリプト情報

- バージョン: v1.1
