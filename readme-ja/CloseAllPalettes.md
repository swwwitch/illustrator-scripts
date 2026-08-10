# CloseAllPalettes

[![Direct](https://img.shields.io/badge/Direct%20Link-CloseAllPalettes.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/CloseAllPalettes.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloseAllPalettes.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

常駐エンジンで動いている各種フローティングパレットをまとめて閉じるユーティリティ。

- 各パレットは `#targetengine` で個別の常駐エンジンに載っており、`$.global` はエンジン
  ごとに独立している。そのため 1 本のスクリプトから他エンジンのパレット参照を直接は
  参照できない
- そこで各エンジンごとに `#targetengine` 付きの BridgeTalk を 1 通ずつ送り、
  受信側エンジンで `$.global.<参照名>` を読んで開いていれば `close()` する
- パレット参照は「Window を直接保持」する形式と「{ window: Window } のラッパー」形式の
  両方に対応する
- 閉じた後は `$.global.<参照名>` を null にして参照を解放する（各パレット本体の onClose
  でも解放されるが保険）
- 対象は PALETTES テーブルで管理。パレットを増やしたら 1 行追加するだけで対象にできる

### 実行時の要点 / Runtime notes

- BridgeTalk のクロスエンジン往復は「ファイル＞スクリプト」実行でも同期的に効く。ただし
  応答が返るまで送信側スクリプトを生かしておく必要があるため、送信ごとに BridgeTalk.pump()
  で応答を待つ（待たずにスクリプトが終わると配信前に終了して閉じ損ねる）
- 送信した BridgeTalk オブジェクトは応答が返るまで配列に保持する（途中で GC されると
  メッセージが配信されず閉じ損ねる）
- 1 件ずつ順に閉じることで、複数エンジンの close() が同一 UI スレッド上で重なって
  ハングするのを防ぐ

### 対象

AiMemoPallete / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManager /
UnifiedTypePanel / ImportAndApplyGraphicStyle / ArtboardDisplayPresetManager /
TextCountStats / SelectionInspector / ApplyLeadingPerTextFrame / TextBreakSplitMergePallete

### スクリプト情報

- バージョン: v1.0.0
