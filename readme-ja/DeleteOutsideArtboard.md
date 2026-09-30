# アートボード外のオブジェクトを削除

[![Direct](https://img.shields.io/badge/Direct%20Link-DeleteOutsideArtboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/DeleteOutsideArtboard.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteOutsideArtboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- ドキュメント内のオブジェクトをアートボードとの重なり条件で判定し、外側のオブジェクトを削除または保管用レイヤーに移動するスクリプトです。
- 「現在のアートボードのみ」または「すべてのアートボード」を対象に選択できます。

![](https://www.dtp-transit.jp/images/ss-500-492-72-20250708-015443.png)

### 主な機能

- アートボード外オブジェクトの削除
- 保管用レイヤーに移動オプション
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログで対象アートボードと移動オプションを選択
2. オブジェクトとアートボードの重なりを判定
3. 重なっていないオブジェクトを削除または保管用レイヤーに移動

### 更新履歴

- v1.0.0 (20250708) : 初期バージョン
- v1.4.2 (20260927) : 「アートボード外：削除」と［保管用レイヤーに移す］を併用すると何もしなかった不具合を修正。ドキュメントが無いときの警告とツールチップを整理
- v1.4.3 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.4.4 (20260928) : ボタン行を共通の部品で組むようにした
- v1.4.5 (20260929) : ダイアログの不透明度を98%に変更
- v1.4.6 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.4.7（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
