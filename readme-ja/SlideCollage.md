# SlideCollage

[![Direct](https://img.shields.io/badge/Direct%20Link-SlideCollage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SlideCollage.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SlideCollage.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 更新日

20260318

### 概要

アクティブなドキュメント上で、指定した .ai / .pdf（PDFはページ指定）をグリッド配置し、ポートフォリオ用のサムネイル一覧を作成します。

・読み込み：アートボード番号（例 1-20 / 1,3,5）を指定して配置
・アイテム：PDFの配置範囲（アート/トリミング/仕上がり/裁ち落とし）を選択、角丸（pt換算）を適用
  - 各アイテムは同サイズの矩形でクリップグループ化し、角丸（ライブエフェクト）はクリップグループに適用

・グリッド：方向（横/縦/ランダム）、列数、間隔を設定
・偶数列：配分モード（偶数列＋1）で偶数列に+1スロットを追加、ずらし（偶数列の上下オフセット）を個別調整
・レイアウト：スケール（自動フィット結果に対する追加倍率）、回転（全体を回転し中心をアートボード中心に合わせる）、位置調整（横/縦）

・マスク：OK時にマージン内側でクリッピング
  - マスク角丸を設定可能（クリップグループに適用）

・アートボード：背景色（HEX指定）を追加可能（マスク対象外）

変更操作の多くはリアルタイムプレビュー（debounce 120ms）で確認できます（読み込みのアートボード番号変更は対象外）。
数値入力欄は ↑↓ / Shift / Option で増減できます。

・内部構造を整理し、main() 内のページ数処理・入力補助・プレビューキャッシュ補助を関数へ分離して保守しやすくしました。

オリジナルアイデア
Slide Collage - Portfolio Layout Generator -
https://slide-collage.vercel.app/

### スクリプト情報

- バージョン: v1.5
