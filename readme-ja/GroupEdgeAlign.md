# 複数のオブジェクトをアートボードの端・中央・ガイドへ整列

[![Direct](https://img.shields.io/badge/Direct%20Link-GroupEdgeAlign.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/GroupEdgeAlign.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlign.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトをひとまとまりとして扱い、アクティブアートボードの端・中央、または条件に合うガイドへ整列するスクリプト。
- グループ化しなくても、選択範囲内の相対位置を保ったまま移動する。
- 整列先は3×3の9点から選択。ダイアログ表示中はプレビューで結果を確認でき、キャンセルで元の位置に戻る。

### 主な機能

- 「整列」パネルの3×3ウィジェットで整列先を指定（左上／上中央／右上／左中央／中央／右中央／左下／下中央／右下）。水平・垂直がまとめて決まる
- 矢印キー（↑↓←→）で1段階ずつ移動。1回押すごとに「スクリプトを1回実行した」のと同じ動きになり、ガイドが複数あるときは押すたびに次のガイドへ進む
- 「ガイドを使用」ONで、アクティブアートボード内側のガイドを整列先にする。縦ガイドは左右、横ガイドは上下に使う
- 「プレビュー境界を使用」で、線幅・効果を含む見た目の境界（ON）と図形本体の幾何境界（OFF）を切り替え。初期値はIllustratorの環境設定「プレビュー境界を使用」に連動
- 「プレビュー」ONで、整列先を変えるたびに結果を画面に反映
- クリッピンググループは、マスク形状（先頭アイテム）の幾何境界で測る
- ファイル名から整列先を自動判定（`GroupEdgeAlignRIGHT.jsx` → 右揃え）
- キーボードショートカット
  - 整列先：W=左上 / E=上中央 / R=右上 / S=左中央 / D=中央 / F=右中央 / X=左下 / C=下中央 / V=右下
  - G=「ガイドを使用」、B=「プレビュー境界を使用」のトグル
- 日本語／英語の自動切り替え

### 使い方

1. 揃えたいオブジェクトを選択します。
2. スクリプトを実行します。
3. 3×3のウィジェットで整列先をクリック（またはショートカットキーを押）します。
4. OKで確定します。

矢印キーだけで動かしたときは、整列先を選ばずにOKするとその位置で確定します。キャンセルすると、矢印キーでの移動も含めて実行前の位置に戻ります。

### オプション

スクリプト冒頭の「ユーザー設定 / User Settings」ブロックで既定の動作を変更できます。

| 変数 | 既定値 | 内容 |
| --- | --- | --- |
| `SHOW_DIALOG` | `true` | `false` にするとダイアログを出さず、ファイル名から判定した方向で即実行 |
| `USE_GUIDES` | `true` | 「ガイドを使用」の初期値 |
| `DEFAULT_ALIGNMENT_SIDE` | `"right"` | ファイル名から整列先を判定できないときの整列先 |
| `GUIDE_SEARCH_MODE` | `"inside"` | `"inside"`＝揃える向きの先にある直近のガイド、`"nearest"`＝向きを問わず最も近いガイド |
| `GUIDE_ORIENTATION_TOLERANCE` | `0.01` | ガイドを水平・垂直と見なす許容値 |

### 注意点

- 対象になるガイドは、アクティブアートボードの内側にあるものだけです。
- 3×3で整列先を選んだときはガイドを使わず、アートボードの端・中央へ揃えます（「ガイドを使用」はディム表示）。ガイドが使われるのは、矢印キーでの移動時と、ダイアログを出さない実行（`SHOW_DIALOG = false`）のときです。
- 中央揃え（上中央／左中央／中央／右中央／下中央）はガイドに吸着しません。
- 方向を固定した単発スクリプトが必要な場合は、`GroupEdgeAlign-7/` の7本（LEFT / RIGHT / TOP / BOTTOM / CENTER / CENTERX / CENTERY）を利用できます。
- ファイル名による判定を使わない版として `GroupEdgeAlignNoFileName.jsx` があります。

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/n4ae0e1e70481)

### 更新履歴

- v1.0 (2025-04-06) : 初版
- v1.0.1 (2026-08-31) : 3×3のラジオボタンを9軸ウィジェット（onDraw描画）に変更、ボタンエリアを左右分割に整理。あわせてヘッダーを概要＋README参照に整理、基本情報に紹介記事URLを追加、ユーザー設定とレイアウトのブロックを分離、LABELSをカテゴリ入れ子に再構成、全関数にJSDocを付与、方向別のif連鎖をテーブルに集約、ガイド探索の重複した判定を統合、プレビューと矢印キーのステップ移動を整列セッションに分離
