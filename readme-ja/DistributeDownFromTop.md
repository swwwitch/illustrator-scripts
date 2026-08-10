# DistributeDownFromTop

[![Direct](https://img.shields.io/badge/Direct%20Link-DistributeDownFromTop.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/DistributeDownFromTop.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeDownFromTop.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択内容に応じて、行送りと配置を次の順で判定して調整します。

1. テキストを 1 つだけ選択しているとき: そのテキストの行送りに「サイズ／行送り」の値を加えます。
2. 複数のテキストを選択していて上端 Y がほぼ同じ（横並び）のとき: 位置は動かさず行送りだけを調整します。行送りがバラバラなら平均値に統一し、揃っていれば全体を「サイズ／行送り」分ずつ増やします。
3. それ以外の複数選択（縦積み）: 最上部のオブジェクトを固定し、以降を「サイズ／行送り」の値ぶんずつ下方向へ等間隔に再配置します。

### 注意点

- 移動・加減算に使う値は、環境設定［テキスト］の「サイズ／行送り」増分（`text/sizeIncrement`）を表示単位込みで pt 換算したものです。
- 行送りは手動行送りではなく、目標行送りから自動行送り量（`autoLeadingAmount` ％）を求めて適用します。
