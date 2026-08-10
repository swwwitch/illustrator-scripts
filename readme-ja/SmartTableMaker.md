# SmartTableMaker

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartTableMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/SmartTableMaker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartTableMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 更新日

20260131

### 概要

2つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、各オブジェクトの背面に左右2分割の背景を作成します。
実行時に高さ倍率（%）を指定するダイアログが表示され、閉じる前にプレビューを確認できます（デフォルトは200%）。

左右の背景幅は、2つのオブジェクト間のギャップを基準に計算されます。
［バランス］パネルで「なし／左／右」を選択すると、
- 「なし」：左右均等（中央分割）
- 「左」　：左側のマージンを［幅］で指定（右側は自動計算）
- 「右」　：右側のマージンを［幅］で指定（左側は自動計算）

［幅］はスライダーおよび数値入力で指定でき、選択中の2オブジェクト間ギャップを最大値として自動設定されます。
テキストのサイドベアリング等によるズレを減らすため、
一時的にテキストをアウトライン化して外接矩形を計算し、計算後すぐに一時生成物を削除します。
その上で背景長方形を選択オブジェクトの背面に配置し、ダイアログ内でリアルタイムにプレビュー表示します（元のオブジェクトは変更しません）。

高さ（%）および［幅］入力欄では、↑↓キーで±1、Shift+↑↓で±10（10刻みスナップ）、Option+↑↓で±0.1 の増減が可能です。

### 更新履歴

- v1.0 (20260124) : 初期バージョン
- v1.1 (20260126) : ［バランス］（なし／左／右）と［幅］指定による左右背景の比率調整に対応。［幅］はオブジェクト間ギャップを最大値として自動計算し、スライダー／数値入力／矢印キー操作に対応
- v1.2 (20260131) : プレビュー時のUndo履歴を汚さないように app.undo() を用いたプレビュー管理（PreviewManager）を導入。OK時は一度ロールバックして本番処理を1回だけ再実行し、Ctrl+Z一回で戻せるように調整

### スクリプト情報

- バージョン: v1.1
