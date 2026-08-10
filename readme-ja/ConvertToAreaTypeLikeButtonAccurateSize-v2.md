# ConvertToAreaTypeLikeButtonAccurateSize-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-ConvertToAreaTypeLikeButtonAccurateSize-v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ConvertToAreaTypeLikeButtonAccurateSize-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButtonAccurateSize-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストを、選択内容に応じて双方向に変換する。

- ポイント文字・パス上文字 → 見た目と実寸を保ったままエリア内文字へ
- エリア内文字 → ポイント文字へ

処理の流れ（ポイント文字・パス上文字 → エリア内文字）：

1. パス上文字はいったんポイント文字へ分離（字形・属性は保持）
2. 変換前に、テキストの見た目をグラフィックスタイルとして一時登録（ダイナミックアクション）
3. 複製→アピアランス分割→アウトラインで計測した実寸の長方形フレームを作り、エリア内文字に変換
4. 元テキストの内容・フォント・サイズ・カーニング・文字組みアキ量設定を引き継ぐ
5. 登録したグラフィックスタイルを適用し、その一時スタイルを削除
6. 行揃え（水平）・テキストの配置（垂直）は常に中央

処理の流れ（エリア内文字 → ポイント文字）：

1. エリア文字の幅・高さを取得（幅A、高さB）
2. convertAreaObjectToPointObject() でポイント文字へ変換
3. 複製→アウトライン化で文字の実寸を計測（幅D、高さE）し、複製は破棄
4. ［形状に変換：長方形］効果を付与。幅に追加 = A − D、高さに追加 = B − E として、元のフレーム実寸の長方形を再現

### 補足

- 選択にエリア内文字が含まれる場合は逆変換（→ ポイント文字）、含まれない場合は順変換（→ エリア内文字）として扱う。
- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

### スクリプト情報

- バージョン: v1.0.0
