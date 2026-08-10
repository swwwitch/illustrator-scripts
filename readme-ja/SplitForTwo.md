# SplitForTwo

[![Direct](https://img.shields.io/badge/Direct%20Link-SplitForTwo.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SplitForTwo.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitForTwo.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 更新日

20260314

### 概要

1つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、そのオブジェクトの外接矩形を左右または上下に2分割し、背面に2色の背景を作成します。
プレビューは専用レイヤーに描画して差し替え、［OK］時に図形や線が二重に作成されないようにします。

［分割方法］で「左右／上下」を選択できます。
［バランス］パネルでは、左・右（または上・下）の幅と比率を数値入力およびスライダーで調整できます。

［線］では外枠と区切り線の有無、線幅、線色を指定できます。
［オプション］では角丸やピル形状を指定できます。
カラー指定は RGB / CMYK / グレーに対応したカラーピッカーから行えます。
プリセットの保存・呼び出し・書き出しにも対応しています。

テキストのサイドベアリング等によるズレを減らすため、
一時的にテキストをアウトライン化して外接矩形を計算し、計算後すぐに一時生成物を削除します。
その上で背景長方形を選択オブジェクトの背面に配置し、ダイアログ内でリアルタイムにプレビュー表示します（元のオブジェクトは変更しません）。

［幅］入力欄では、↑↓キーで±1、Shift+↑↓で±10（10刻みスナップ）、Option+↑↓で±0.1 の増減が可能です。

### 更新履歴

- v2.9.2 (20260314) : バージョン表記と更新日を更新。

### スクリプト情報

- バージョン: v2.9.2
