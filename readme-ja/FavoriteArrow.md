# よく使う矢印と線の設定をまとめて適用

[![Direct](https://img.shields.io/badge/Direct%20Link-FavoriteArrow.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/FavoriteArrow.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteArrow.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したパスに、よく使う矢印と線の設定（線幅・線端・角の形状・破線）をまとめて適用します。
- 矢印は Illustrator の DOM から操作できないため、一時アクション（ai_plugin_setStroke）を生成して実行します。
- [SetStrokeAndArrowheads](SetStrokeAndArrowheads.md) をもとに、SetStrokeAlignment の線端・角の形状と、[DashGapCalculator](DashGapCalculator.md) の破線の計算を取り込んだ改編版です。

### 主な機能

#### 線

- 線幅
- 線端（なし／丸型／突出）、角の形状（マイター／ラウンド／ベベル）

#### 矢印

- よく使う矢印（矢印 1・8・11）をラジオボタンで選択。選ぶと倍率がその矢印に合わせた値に変わります（1・11 は 100%、8 は 25%）
- それ以外の矢印はポップアップメニューから選択（倍率は 100%）
- ［終点も同じ］：終点にも始点と同じ矢印を付けます。オフのときは終点の矢印を外します
- ［始点と終点を入れ替え］：矢印を始点ではなく終点に付けます（［終点も同じ］がオンのときはディム）
- 先端位置（パスの終点に配置／パスの終点から配置）

#### 破線

- ［破線］［ドット点線］のどちらか一方を選択。オンにすると、線幅から決めた分割数・間隔・線分が入ります
- 破線の計算：分割数・間隔・線分から破線を求めます。複数のパスを選んでいるときは、パスごとに長さを見て計算します
- 計算方法：間隔→線分／線分→間隔
- ドット点線は点（長さ0）の間隔を分割数から求め、線端を丸型に固定します
- ［両端を調整］：オープンパスで、両端が線分（点）で終わるように配分します

#### プレビュー

- 線幅は入力中も即時に反映、矢印を含む設定はアクション実行＋取り消しで反映します。キャンセルすると元に戻ります

### 使い方

1. 線を設定したいパスを選択します（グループ・複合パスの中のパスも対象）
2. スクリプトを実行し、ダイアログで線・矢印・破線・先端位置を設定します
3. ［OK］で適用します

### 注意点

- 矢印名・先端位置名・線端と角の形状の名前は、Illustrator の UI 表示名と一致している必要があります（言語に依存）。
- 矢印の倍率キー（asc1 / asc2）は推定値です。
- 破線はアクションの後に DOM で設定しています。
- よく使う矢印と倍率は、スクリプト冒頭の `FAVORITE_ARROWS` で変更できます。

---

### 更新履歴

- v1.0.0（2026-10-03）初版
- v1.0.1（2026-10-04）線幅の初期値を一般の単位に合わせるようにした（mm は 0.25pt、px は 1px、それ以外は 5pt）

### スクリプト情報

- バージョン: v1.0.1
- 初回リリース: 2026-10-03
- 最終更新: 2026-10-04
