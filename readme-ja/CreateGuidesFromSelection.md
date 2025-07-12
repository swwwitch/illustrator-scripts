# 選択したオブジェクトに対してガイドを自動作成

[![Direct](https://img.shields.io/badge/Direct%20Link-CreateGuidesFromSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CreateGuidesFromSelection.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- Illustrator の選択オブジェクトからガイドを作成するスクリプト。
- ダイアログ上で上下左右中央のガイドを自由に指定して描画可能。

![](https://www.dtp-transit.jp/images/ss-588-938-72-20250712-205442.png)

### 主な機能

- プレビュー境界または幾何境界の選択
- 一時アウトライン化とアピアランス分割による正確なテキスト処理
- オフセットと裁ち落とし指定
- 「_guide」レイヤー管理とガイド削除オプション
- クリップグループと複数選択対応

### 処理の流れ

- ダイアログボックスでオプションを設定
- 選択オブジェクトの外接矩形を取得
- ガイドを描画
- テキストオブジェクトは一時アウトライン化およびアピアランス展開して計算後復元
- 「_guide」レイヤーにガイドを追加後ロック

### note

https://note.com/dtp_tranist/n/nd1359cf41a2c

### 更新履歴


- v1.0 (20250711) : 初期バージョン
- v1.1 (20250711) : 複数選択、クリップグループ、日英対応追加
- v1.2 (20250711) : プレビュー境界OFF時の安定化、テキスト処理修正
- v1.3 (20250711) : アートボード外のオブジェクト自動カンバス選択、テキストアウトライン処理改善
- v1.4 (20250711) : アートボード外のオブジェクト選択時のアラート削除、カンバス選択時のはみだし無効化
- v1.5 (20250711) : 左上、左下、右上、右下モードを追加（それぞれ2本のガイドを作成）
- v1.6 (20250712) : コードリファクタリング、ラジオボタンの表示切り替え機能追加
- v1.6.1 (20250712) : 微調整
- v1.6.2 (20250712) : 単位設定