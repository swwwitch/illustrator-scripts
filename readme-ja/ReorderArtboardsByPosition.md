# カンバス上の並びで［アートボード］パネルの並び順を変更

[![Direct](https://img.shields.io/badge/Direct%20Link-ReorderArtboardsByPosition.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ReorderArtboardsByPosition.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- アートボードを名前順や位置順（左上、右上、上左、上右）で並べ替えるスクリプト
- カンバス（ドキュメント）上の見た目の並びを基に、［アートボード］パネルの順序を変更

![](https://www.dtp-transit.jp/images/ss-544-634-72-20250707-032437.png)

### 主な機能

- アートボードを名前順または位置順にソート
- 許容差（Tolerance）の設定により微妙なズレを許容
- UI でソート方法と許容差を選択可能

### 処理の流れ

1. ダイアログでソート方法と許容差を選択
2. OK ボタンで並べ替えを実行
3. 結果を即時反映

### オリジナル、謝辞

- m1b 氏: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### note

- https://note.com/dtp_tranist/n/nb416cb01728a

### 更新履歴

- v1.0.0 (20231115) : 初期バージョン（Andrew_BJ による UI 改良と上限拡張）
- v1.1.0 (20231116) : 許容差の自動計算機能とスライダーを追加、ロジック整理
- v1.2.0 (20260415) : 左上基準専用ツールとしてUIと構成を整理、再配置設定とプレビュー表示を調整
- v1.3.0 (20260508) : アートボード名パネル追加（「行-列」形式の自動命名、既存名の整形、区切り文字／桁数の選択）、再配置ロジックとUI構築を責務ごとの関数に分割、ローカライズを1つの関数に統一、列間／行間の連動処理を整理
- v1.3.1 (20260508) : 再配置とパネル並び順の実行順を入れ替え、再配置後の見た目に合わせてパネル順が更新されるよう修正
- v1.4.0 (20260513) : パネル上の並び順を「名前順／カンバス上の並び順に／変更しない」のラジオボタンに変更、名前順は数字列を10桁ゼロ埋めした自然順ソート
- (20260807) : 概要と基本情報ブロックを共通書式に統一、レイアウト定数とUIヘルパーを整理、全関数にJSDocを付与（機能変更なし）