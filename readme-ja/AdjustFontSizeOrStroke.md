# テキストは文字サイズ、それ以外は線幅を一定量ずつ増減

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustFontSizeOrStroke.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AdjustFontSizeOrStroke.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSizeOrStroke.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストは文字サイズを、それ以外のオブジェクトは線幅を、環境設定の単位で一定量ずつ増減します。ダイアログボックスは出さずにすぐ実行するので、キーボードショートカットに割り当てて使うことを想定しています。

### 主な機能

- テキストは文字サイズを、環境設定の［単位］の［文字］で 1 増やす（ポイントなら 1 pt、級なら 1 Q）
  - 文字ごとにサイズが異なるテキストは、それぞれのサイズに足す
  - 文字ツールで文字を選択しているときは、その文字だけを変える
- テキスト以外は線幅を、環境設定の［単位］の［線］で 0.1 増やす（ポイントなら 0.1 pt、ミリメートルなら 0.1 mm）
  - 線が設定されていないパスは変えない
- グループ・複合パスは中までたどり、テキストとパスが混ざっていてもそれぞれを変える
- 足したあとは単位の値で丸めるので、繰り返し実行しても端数が残らない

### オプション

スクリプト冒頭の「ユーザー設定」で切り替えます。

| 変数 | 値 | 内容 |
|---|---|---|
| `STEP_SIGN` | `1` / `-1` | `1` で大きく・太く、`-1` で小さく・細く |
| `STEP_MULTIPLIER` | `1` / `5` / `10` | 1回に変える量の倍率（×1・×5・×10） |
| `FONT_SIZE_STEP` | `1` | 文字サイズの1回分の量（［文字］の単位） |
| `STROKE_STEP` | `0.1` | 線幅の1回分の量（［線］の単位） |

増やす・減らす、倍率ごとにショートカットを分けたいときは、ファイルを複製して設定を変えたものをそれぞれ割り当てます。

### 注意点

- 文字サイズは 0.1〜1296 pt、線幅は 0.1〜1000 pt の範囲で止まります（細くしても線が見えなくならないよう、0 pt にはしません）。
- テキストの文字の線幅は変えません（テキストは文字サイズだけを変えます）。
- アピアランスで追加した線は変えません（基本の線だけを変えます）。
- グループの中のロック・非表示のオブジェクトは変えません。
- 行送りが自動でないテキストは、文字サイズを大きくすると行が詰まって見えることがあります。

### 更新履歴

- v1.0.0 (2026-11-01) : 初版
- v1.0.1 (2026-11-01) : 線幅の下限を 0 pt から 0.1 pt に変更（細くしても線が「あり」のまま見えなくならないように）
