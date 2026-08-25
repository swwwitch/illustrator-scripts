# ExcelTableNormalizer

[![Direct](https://img.shields.io/badge/Direct%20Link-ExcelTableNormalizer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/ExcelTableNormalizer.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExcelTableNormalizer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

Excel由来のIllustratorデータを、表組みとして扱いやすい状態に整形します。

### 主な機能

- クリッピングマスクを解除し、エラーインジケーターなどの小さな不要オブジェクトを削除
- テキストを専用レイヤー（`_text_all`）へ移動し、重複テキストを削除
- 列ごとの配置をダイアログで指定（左／中央／右、自動判定あり）
- テキストを印刷用の黒へ変換
- 罫線グリッドから列を推定し、列サンプルを表示
- セル背景を専用レイヤー（`_cell_rectangle`）へ抽出・調整し、必要に応じて高さを均等化
- 長方形状の罫線を中心線化し、均等配置・結合セル対応・列幅保持を実行

### 使い方

1. Excelから貼り込んだ表のオブジェクトを含むドキュメントを開きます。
2. スクリプトを実行します。
3. 列ごとの配置などを指定して実行します。

### 注意点

- 実行時にカラーモード関連のメニューコマンドを実行し、以後の色指定を安定させます。
- 元データを大きく作り替えるため、実行前にファイルを複製しておくことを推奨します。

### 更新履歴

- v1.1.0 (2026-04-30)
