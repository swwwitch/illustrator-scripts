# TextWithShapeToAreaTypeSimple

[![Direct](https://img.shields.io/badge/Direct%20Link-TextWithShapeToAreaTypeSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextWithShapeToAreaTypeSimple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextWithShapeToAreaTypeSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ポイント文字・パス上文字を、見た目（アピアランス）を保ったままエリア内文字へ変換する。

処理の流れ：

1. 実行時にオプションダイアログを表示（大きさ調整の する/しない と 幅%・高さ%）
2. 変換前に、選択テキストの見た目を一時グラフィックスタイルとして登録
3. ポイント文字 / パス上文字 → 計測した実寸（大きさ調整ONなら幅・高さに倍率）でエリア内文字へ変換し、登録した一時スタイルを適用してアピアランスを引き継ぐ（適用後に一時スタイルは削除）
4. 大きさ調整ONのときはボタン背景（New Fill を2枚追加＋長方形シェイプ効果）を付与
5. 行揃え・テキストの配置は常に中央

### 補足

- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。
- 1件も変換できなかった場合は無言で終わらず警告を表示する。

### 対応している選択

- ポイント文字 … そのまま計測してエリア内文字化
- パス上文字 … 内部で一度ポイント文字に分離（detachPathTextToPointText）してから同じ経路で処理
- 複数選択時はそれぞれ個別に変換

### オプションの効き方

| | 大きさ調整の倍率 | ボタン背景（塗り2枚＋長方形効果） |
|---|---|---|
| する | ○（幅×1.2 / 高さ×1.6 等） | ○ |
| しない | 等倍 | ✕ |

- ボタン背景は New Fill を2枚追加し、長方形シェイプ効果で背景化する。塗り色は設定しない（追加した塗りは既定のまま）。

### 更新履歴

- v1.3.0 (2026-07-02): 「スタイルの読み込み」機能（読み込み／再読み込みボタン、外部 AI ファイルからのスタイル取り込み、参照ファイルの記憶）、グラフィックスタイルの選択 UI（元の見た目／読み込んだスタイルのラジオ）、テキスト＋長方形の合体ロジック（長方形をフレーム化してテキストを流し込む処理）を撤去。ポイント文字・パス上文字を実寸計測でエリア内文字へ変換する処理に単純化。アピアランスの引き継ぎ（選択テキストの見た目を一時グラフィックスタイルとして登録・適用し、適用後に削除）は従来どおり維持 ／ Removed the "Load Styles" feature (Load/Reload buttons, importing styles from an external AI file, remembering the source file), the graphic-style selection UI (original/loaded-style radios), and the text + rectangle merge logic (turning a rectangle into the frame and flowing the text in). Simplified to converting point/path text into area type at the measured real size, while keeping the appearance inheritance as before (register the source text's look as a temp graphic style, apply it, then remove it)
- v1.2.0 (2026-07-02 追記): グラフィックスタイルのラジオ（元の見た目／読み込んだスタイル）を排他選択に修正（別コンテナのため自動排他が効いていなかった）。ダイアログで読み込み／再読み込みした後に変換すると選択が復帰されず無言で何も起きない問題を修正（選択復帰を変換直前に常時実行）。再読み込み後にどのラジオも未選択になる問題を修正（同名復元／なければ元の見た目へ）。変換が0件のとき無言終了せず、握り潰していた例外理由を添えて警告を表示 ／ Made the graphic-style radios (original appearance / loaded styles) mutually exclusive (they lived in separate containers, so ScriptUI's auto-exclusion didn't apply); fixed a silent no-op when converting after a Load/Reload (selection is now always restored right before converting); fixed a no-selection state after Reload (restore by name, else fall back to original); a 0-result conversion now alerts with the previously swallowed error instead of exiting silently
- v1.2.0 (2026-07-01): 固定パス（TARGET_FILE_PATH）と固定スタイル名（文字白抜き／枠のみ）を撤去。ダイアログに「スタイルの読み込み」ボタンを追加し、選んだ AI ファイルとスタイル名を Folder.userData に記憶。取り込んだスタイル名からラジオを自動生成する構成へ変更（ImportGraphicStyles v1.7.0 より移植）／ Removed the hardcoded path (TARGET_FILE_PATH) and fixed style names (white text / frame only); added a "Load Styles" button that remembers the picked AI file and style names in Folder.userData; radios are now generated automatically from the imported style names (ported from ImportGraphicStyles v1.7.0)
- v1.1.0 (2026-07-01): ダイアログにグラフィックスタイル選択（元の見た目／文字白抜き／枠のみ）を追加。外部 AI ファイル（TARGET_FILE_PATH）から未登録スタイルを取り込んで適用（ImportGraphicStyles より移植）。角丸オプションを削除／ Added a graphic-style choice (original appearance / white text / frame only); imports the chosen style from an external AI file when it is not yet registered (ported from ImportGraphicStyles). Removed the round-corners option
- v1.0.0 : 初期バージョン / Initial release

### スクリプト情報

- バージョン: v1.3.0
