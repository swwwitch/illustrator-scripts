# ガイドを線付きのパスに変換してコピー

[![Direct](https://img.shields.io/badge/Direct%20Link-CopyGuidesAsPaths.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CopyGuidesAsPaths.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyGuidesAsPaths.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

アクティブアートボード内のガイドを、線の付いた通常のパスに変換してクリップボードへ送ります。元のガイドはドキュメントに残るため、ガイドを保ったまま別のドキュメントやアプリケーションへ罫線として持ち出せます。

### 主な機能

- アクティブアートボード内のガイドだけを対象（ガイドの中心点で判定）
- ロックされたオブジェクト、ロックされたレイヤー上のガイドも対象
- 同じ座標に重なったガイドは1本だけ残し、余分なものはドキュメントから削除
- 直線だけでなく、曲線やクローズパスのガイドも形状そのままで変換
- 変換したパスは作業用レイヤー上で作ってカットするため、ドキュメントには残らない
- 線色はドキュメントのカラーモードに応じて K100（CMYK）／黒（RGB）

### 使い方

1. ガイドのあるアートボードをアクティブにします。
2. スクリプトを実行します。
3. 「n本のガイドをコピーしました。」と表示されたら、貼り付け先でペーストします。

### 処理の流れ

- アクティブアートボード内のガイドを収集
- 同じ座標に重なったガイドを削除
- 作業用レイヤーを作成し、各ガイドのアンカーと方向線を写して線付きパスを作成
- 作成したパスだけを選択してクリップボードへカット
- 作業用レイヤーを削除し、アクティブレイヤーを元に戻す

### オプション

スクリプト冒頭の「ユーザー設定 / User Settings」ブロックで変更できます。

| 変数 | 既定値 | 内容 |
| --- | --- | --- |
| `STROKE_WIDTH` | `0.5` | 変換後のパスに付ける線幅（pt） |
| `MATCH_TOLERANCE` | `0.001` | 同じ座標とみなす誤差（pt） |
| `TEMP_LAYER_NAME` | `"__guide_to_path__"` | 作業用レイヤー名 |

### 注意点

- 非表示のガイド、非表示レイヤー上のガイドは対象外です。
- アートボードの判定はガイドの中心点で行うため、アートボードをまたぐ長いガイドで中心が範囲外にあるものは対象外になります。
- 重複ガイドの削除はドキュメントを変更します（取り消し可能）。
- 環境設定「ペースト時にレイヤーを保持」がONの場合、ペースト先に `__guide_to_path__` レイヤーが作られます。
- 吸着させたガイドは座標が 1e-12 単位でずれることがあるため、重複判定には `MATCH_TOLERANCE` の許容値を使っています。

### 更新履歴

- v1.0.0 (20260907) : 初期バージョン
