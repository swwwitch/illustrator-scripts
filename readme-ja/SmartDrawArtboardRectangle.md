# SmartDrawArtboardRectangle

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartDrawArtboardRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/SmartDrawArtboardRectangle.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDrawArtboardRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- アクティブまたは全アートボードと同サイズの長方形を、オフセットを考慮して描画します。
- カラー（なし／K100 15%／HEX／CMYK）、重ね順（最前面／最背面／bgレイヤー）を指定でき、ライブプレビューで確認できます。
- 描画後に「ガイド化」「ライブシェイプに変換」をオプションで適用できます。
- ライブシェイプには中心の○（中心点）を常に表示します。

### 主な機能

- オフセット指定（裁ち落としプリセット：3mm／12H／0.125in）
- カラー指定（None／K100 15%／HEX／CMYK）
- 重ね順（Front／Back／bgレイヤー、デフォルトは最前面）
- 対象範囲（作業アートボード／すべて）
- オプション（ガイド化：デフォルトOFF／ライブシェイプに変換：デフォルトON）
- 中心の○（中心点）を常に表示（記録済みアクションで適用）
- プレビュー（1ptの破線、50%トーン、専用レイヤー）
- ホットキー（F：最前面／B：最背面／L：bgレイヤー／G：ガイド化／C：現在のアートボード／A：すべて）
- ダイアログの初期位置・不透明度の設定

### 紹介記事

https://note.com/dtp_tranist/n/n1ba88513a9c8

### 更新履歴

- v1.0 (20250820) : 初期バージョン
- v1.5.1 (20250824) : ライブシェイプ化・ガイド化などのオプション追加
- v1.5.2 (20260531) : 中心の○を常に表示、単位テーブル統合、CMYK入力修正、ホットキー再編
- v1.5.3 (20260531) : オブジェクト名を「<長方形>」に変更、オフセット入力欄の幅を調整
- v1.5.4 (20260601) : 最前面がテンプレート／ロックレイヤーのときの Error 8705 を修正
- v1.5.5 (20260625) : プレビューレイヤーの後始末を確実化（残存・誤描画を防止）、ガイド化時はオブジェクト名を「<ガイド>」にして選択解除、HEX欄で色名／#RGB短縮／grayNN を許容、ホットキー抑止の堅牢化、中心の○はライブシェイプ時のみ、内部リファクタリング、最新バージョン

---

### スクリプト情報

- バージョン: v1.5.5
