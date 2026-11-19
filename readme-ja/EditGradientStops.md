# グラデーションストップを一覧表示して編集

[![Direct](https://img.shields.io/badge/Direct%20Link-EditGradientStops.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/EditGradientStops.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditGradientStops.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトに使われているグラデーションのストップを一覧表示し、カラー・位置・不透明度の変更や、ストップどうしの入れ替えができます。

### 主な機能

- 塗りと線のグラデーションを、グループの中・複合パス・テキストまで集めて、グラデーションごとのパネルに一覧表示
- 各ストップをカラーの見本・カラーの値（CMYK・RGB・グレー・Lab・スウォッチ名と濃度）・位置・不透明度の1行で表示（位置の順）
- 見本をクリックすると Illustrator 標準のカラーピッカーでカラーを変更
- 位置・不透明度は ∧∨ ボタン・↑↓キー・直接入力（四則演算も可）で変更（0〜100%）
- 行の末尾の ↑↓ ボタンで、上下のストップとカラー・不透明度を入れ替え（位置はそのまま）
- 行の末尾の複製ボタンで、ストップを複製してすぐ下に追加（位置は次のストップとの中間、最後の行なら前のストップとの中間）
- 行の末尾の削除ボタンで、ストップを削除（2つより少なくはできない）
- ［プレビュー］（option＋A、最初からオン）で、変更をその場でオブジェクトに反映
- ストップの構成が同じグラデーションは1つのパネルにまとめ、変更はそれを使うすべての塗り・線に適用

### 使い方

1. グラデーションを使ったオブジェクトを選択します。
2. スクリプトを実行します。
3. カラー・位置・不透明度を変更し、［OK］をクリックします。［キャンセル］では元に戻ります。

### 注意点

- アピアランスパネルで追加した塗り・線のグラデーションは対象外です。
- テキストは、テキストフレーム全体の文字属性として返る塗り・線を読み書きします。
- カラーピッカーはドキュメントのカラーモードの色を返すため、スウォッチのカラーを選び直すとプロセスカラーになります。
- 一覧に表示するグラデーションは10個までです（スクリプト冒頭の MAX_LISTED_GRADIENTS で変更）。
- ［プレビュー］を最初からオンにするかは、スクリプト冒頭の DEFAULT_PREVIEW_ON で変えられます。

### 紹介記事

[【Illustrator】グラデーションストップを一覧表示して編集する｜DTP Transit 別館](https://note.com/dtp_tranist/n/n524fafa198dd)

### 更新履歴

- v1.0.0（2026-11-20）最初のリリース
