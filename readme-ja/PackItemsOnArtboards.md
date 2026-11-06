# オブジェクトをアートボードに詰めて配置

[![Direct](https://img.shields.io/badge/Direct%20Link-PackItemsOnArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/PackItemsOnArtboards.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PackItemsOnArtboards.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトを、すき間を詰めてアートボードに配置します。大きさの異なるオブジェクトを、並べ替えの方法を変えながら何度も配置してみて、最もよい結果を採用します。間隔・余白・回転の可否・試行回数はダイアログボックスで指定できます。

### 主な機能

- アートボードの順に詰めて配置します。1枚目に入りきらない分は次のアートボードへ送ります。
- ［オブジェクトの間隔］と［アートボードの余白］を指定できます（初期状態は 1 mm と 5 mm）。どちらもマイナスの値を指定でき、［オブジェクトの間隔］をマイナスにすると重ねて配置します。
- ［最大試行回数］で、配置を試す回数の上限を指定します。開いたときの値はオブジェクトの数から計算します。
- ［優先］で、試行の結果を比べるときに重視する値を［配置した個数］か［配置した面積］から選びます。
- 試行のたびに、面積・長辺・幅・高さ・大小交互の順に並べ替えを変え、そのあとはランダムに並べます。
- ［最大試行回数まで試す］をオンにすると、すべて配置できても上限まで、よりよい結果を探します。
- ［重ね順のまま配置］をオンにすると、大きさで並べ替えず、重ね順のまま1回だけ配置します。
- ［90°回転を許可］をオンにすると、入りきらないオブジェクトを90°回転して配置します。
- ［任意の角度の回転を許可］をオンにすると、配置の前に、外接する長方形が最小になる角度へ回転します。
- ［結果を表示］をオンにすると、配置のあとに、採用した試行とアートボードごとの個数を表示します（初期状態はオフ）。
- ［ランダムに試す］で、ダイアログボックスを閉じずに、ランダムな順で1回配置できます。
- 配置できなかったオブジェクトは、選択された状態で残ります。

### 使い方

1. オブジェクトを選択します。
2. スクリプトを実行します。
3. 間隔・試行・回転を指定します。
4. ［実行］をクリックします。

### 注意点

- グループは1つのオブジェクトとして動かします。
- 大きさは線幅を含む見た目の範囲で測ります。ポイント文字はアウトラインの形、クリップグループはマスクの内側の範囲です。
- オブジェクトは外接する長方形で詰めるので、形のくぼみには入り込みません。
- ［ランダムに試す］のあとに［キャンセル］で閉じても、配置は残ります。
- 試行回数やオブジェクトの数が多いと、時間がかかることがあります。

### クレジット

m1b さんの [Bin Packing For Illustrator and Indesign](https://github.com/mark1bean/bin-packing-for-illustrator-and-indesign)（MIT License）を改変したものです。配置の計算には、trentium さんが [Stack Overflow](https://stackoverflow.com/questions/56642111/bin-packing-js-implementation-using-box-rotation-for-best-fit) で公開したコードをもとにした処理を含みます。

### 更新履歴

- v1.2.0 (2026-11-07) : 公開。［結果を表示］の初期状態をオフに変更。ライブラリを本体に取り込み、このスクリプトだけで動くように変更
