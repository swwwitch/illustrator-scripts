# ConvertToAreaTypeLikeButtonAccurateSize

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToAreaTypeLikeButtonAccurateSize.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertToAreaTypeLikeButtonAccurateSize.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButtonAccurateSize.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する。

処理の流れ：

1. 変換前に、選択テキストの見た目をグラフィックスタイルとして一時登録（ダイナミックアクション）
2. ポイント文字 / パス上文字 → 計測した実寸のエリア内文字へ変換
   テキスト＋図形 → 図形をエリア内文字にしてテキストを流し込む
3. 変換後のエリア内文字に登録したグラフィックスタイルを適用し、その一時スタイルを削除
4. 行揃え・テキストの配置は常に中央

### 補足

- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

### スクリプト情報

- バージョン: v1.0.0
