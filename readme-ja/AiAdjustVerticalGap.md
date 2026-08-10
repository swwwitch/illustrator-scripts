# AiAdjustVerticalGap

[![Direct](https://img.shields.io/badge/Direct%20Link-AiAdjustVerticalGap.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/AiAdjustVerticalGap.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAdjustVerticalGap.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです。
ライブプレビュー対応で、設定を変えるたびに結果を確認できます。

- 対象は「2つのオブジェクト選択」または「2点を含むグループ1つの選択」です
- 起動時に選択2点の現在の間隔を読み取り、間隔値に取り込みます（その場では動きません）
- 「キーオブジェクト」（上／下）を基準に、もう一方を移動します
- 間隔値は定規の単位で指定でき、↑↓キー（Shiftで±10／Optionで±0.1）で増減できます
- 間隔値に負の値を指定すると、2点を重ねられます（オーバーラップ）
- 左右方向の整列（なし／左／中央／右）も同時に行えます
- 整列後にさらに左右へずらす「横調整」値を指定できます（正の値で右・負の値で左、単位は定規に従う）。整列「なし」でもずらしのみ適用できます
- テキストには段落の行揃え（変更しない／整列に連動／均等配置）を適用できます。左揃えは Illustrator のバグを resize で回避して実現します
- クリップグループはクリップパスを基準に、プレビュー境界（線幅・効果込み）の使用も切替可能
- ［記録］で現在の設定を記録してパネルをロック（ディム表示）し、ボタンは［編集］に切り替わります（再クリックでロック解除）
- ロック中に複数のグループを選択して［適用］すれば、記録した設定をまとめて一括適用できます（各グループの2点／2点選択が対象）
- 未確定のプレビューのまま閉じると元に戻ります
- キー操作：T＝上／B＝下、N／L／C／R＝整列しない／左／中央／右、S＝整列に連動／J＝均等配置、A＝適用、Esc＝閉じる（ロック中はパネル系ショートカットは無効）

### 実装メモ / Implementation note

常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理（runAdjustment と
その補助関数）はメインエンジンへ BridgeTalk で都度委譲する。委譲本文は関数を toString()
で連結し、encodeURIComponent + eval(decodeURIComponent()) で送って文字化けを防ぐ。
委譲する関数では行コメントを使わずブロックコメントのみ、必ずセミコロンで終える。
ライブプレビューは各プレビューを1トランザクションとし、次回送信時に worker 側で
app.undo() してから再適用する（確定時は取り消さない）。

### スクリプト情報

- バージョン: v1.3.0
