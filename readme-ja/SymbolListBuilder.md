# SymbolListBuilder

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolListBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolListBuilder.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolListBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

Illustrator ドキュメントに登録されたシンボルを一覧表示する専用アートボード「シンボル一覧」を自動生成するスクリプト。ダイアログでパラメータを操作しながらライブプレビューでき、OK で確定（プレビュー削除 → 最終ビルド → 旧版掃除）。

### 主な機能

- 作成位置の基準：最終アートボード／番号指定（既定はキャンバス最右下のアートボードを自動採用）
- 作成方向：基準アートボードの右側／下側を選択
- サイズと余白：幅・高さ・内側余白（幅変更時は最大幅も自動追従）
- 背景色：なし／黒／白／グレー（K50）。背景黒のときキャプションを白に
- シンボルの絞り込み：すべて／使用中のみ
- キャプション：しない／上／下、フォントサイズ（Illustrator の文字設定単位）
- 既定キャプションフォントはロケール別（ja → HiraginoSans-W3 / en → MyriadPro-Regular）
- レイヤー／アートボード名はロケール別（ja「シンボル一覧」／ en「Symbol List」）。既存判定は両言語に対応
- 「更新」ON で既存「シンボル一覧」アートボードと、その上に乗っているオブジェクトをすべて削除して置換

### 単位系 / Units

- ルーラー単位 (rulerType) … 寸法・マージン・間隔
- テキスト単位 (text/units) … フォントサイズ

### 紹介記事（note）

https://note.com/dtp_tranist/n/ncac687d0a3a0

### 更新履歴

- v1.0.0（2026-05-09）：初版 / Initial release.
- v1.2.1（2026-06-03）：作成位置を「基準（最終アートボード／番号指定）＋方向（右側／下側）」の 2 グループに再編し、指定番号の既定値にキャンバス最右下のアートボードを自動採用。更新 ON 時は旧「シンボル一覧」アートボード上のオブジェクトをすべて削除。レイヤー／アートボード名をロケール別にし、既存判定は両言語対応。
- v1.2.2（2026-06-03）：パネル名・ラベル・ツールチップの文言を整理。画面ズームスライダーを廃止し、ビュー合わせボタン（シンボル一覧＝作成アートボードにフィット＋90%／全体表示＝全アートボードにフィット＋90%）を最下段ボタン行の左に配置。更新チェックを「作成するアートボード」パネル末尾へ移動。収集対象ラジオを横並びに。

### スクリプト情報

- バージョン: v1.2.2
