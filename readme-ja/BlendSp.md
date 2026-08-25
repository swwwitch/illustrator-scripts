# BlendSp

[![Direct](https://img.shields.io/badge/Direct%20Link-%20BlendSp.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/blend/%20BlendSp.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendSp.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択内容に応じて、ブレンドの作成・設定・調整を1つのダイアログで行うスクリプトです。

選択にブレンドが含まれていれば設定と調整だけを行い、含まれていなければブレンドを作成してから設定します。

### 主な機能

- ステップ数の指定（0〜1000の整数）
- ステップ数スライダー（通常0〜32、option併用で0〜128、shift併用で0〜1000）
- 方向の指定と反転
- 解除／拡張／ブレンド軸を置き換え
- ステップ数と方向のライブプレビュー（Undoで戻せる方式）
- 日本語／英語UI

### 使い方

1. ブレンド、またはブレンドのもとになるオブジェクトを選択します。
2. スクリプトを実行します。
3. ステップ数と方向をプレビューを見ながら調整します。
4. ［OK］で確定します。

### 注意点

- ブレンドを選択して開いた場合、そのステップ数を読み取って初期値にします（取得できない場合は8）。ブレンドを選択していない場合の初期値も8です。
- 「解除」「拡張」「ブレンド軸を置き換え」は安全のためプレビューでは実行せず、［OK］を押したときにだけ実行します。
- 「その他」が「なし」以外のときは、方向／反転パネルが無効になります（状態は保持されます）。
- ブレンド内のパスを選択している場合もブレンド本体を探して同様に扱いますが、環境による差があります。

### 更新履歴

- v1.2 (2026-01-01)
