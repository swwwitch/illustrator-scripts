# ImportAndApplyGraphicStyle

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportAndApplyGraphicStyle.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ImportAndApplyGraphicStyle.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportAndApplyGraphicStyle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する。

処理の流れ：

1. 実行時に常駐パレットを表示（大きさ調整の する/しない と 幅%・高さ%、グラフィックスタイル：元の見た目／読み込んだスタイル）。「変換」ボタンで実行し、パレットは開いたまま
2. 変換前に、適用するグラフィックスタイルを用意（「元の見た目」＝選択テキストの見た目を一時登録／読み込んだスタイル＝「読み込み」で選んだ AI ファイルから現書類へ取り込み）
3. ポイント文字 / パス上文字 → 計測した実寸（大きさ調整ONなら幅・高さに倍率）でエリア内文字へ変換
   テキスト＋長方形 → 長方形を複製してエリア内文字にし、テキストを流し込む
4. ボタン背景（テキストのみ・する ／ テキスト＋長方形。※外部スタイル未使用時のみ）→ New Fill を2枚追加＋長方形シェイプ効果
5. 変換後のエリア内文字に用意したグラフィックスタイルを適用（「元の見た目」の一時スタイルのみ削除し、取り込んだ外部スタイルは残す）
6. 行揃え・テキストの配置は常に中央

### 補足

- テキスト＋図形のフレームに使える図形は長方形のみ（軸並行・直線コーナーの長方形）。長方形以外の閉じたパスのみが図形として選ばれている場合は警告して中止する。
- テキスト＋図形は 1 組（最初のテキスト＋最初の長方形）のみ処理する。
- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

### 対応している選択

① テキストのみ
- ポイント文字 … そのまま計測してエリア内文字化
- パス上文字 … 内部で一度ポイント文字に分離（workerDetachPathText）してから同じ経路で処理
- 複数選択時はそれぞれ個別に変換

② テキスト＋長方形
- テキスト（ポイント文字／パス上文字）＋ 長方形（軸並行・直線コーナー）
- 長方形を複製してエリア内文字のフレームにし、テキストを流し込む
- 1組（最初のテキスト＋最初の長方形）のみ処理

### オプションの効き方

| | 大きさ調整の倍率 | ボタン背景（塗り2枚＋長方形効果） |
|---|---|---|
| テキストのみ・する | ○（幅×1.2 / 高さ×1.6 等） | ○ |
| テキストのみ・しない | 等倍 | ✕ |
| テキスト＋長方形 | ―（長方形サイズ優先） | ○ |

- ボタン背景は New Fill を2枚追加し、長方形シェイプ効果で背景化する。塗り色は設定しない（追加した塗りは既定のまま）。
- グラフィックスタイルで「読み込んだスタイル」を選んだ場合、そのスタイルが見た目を定義するためボタン背景は付与しない。現書類に未登録なら、記憶した AI ファイルから取り込んで適用する。参照ファイルとスタイル名は Folder.userData（styles_for_TextWithShapeToAreaType.txt）に記憶する。

### 更新履歴

- v1.3.0 (2026-07-01): モーダルダイアログを常駐パレット化（#targetengine ＋ $.global 単一インスタンスガード）。DOM 処理は BridgeTalk でメインエンジンへ委譲し、パレットは開いたまま実行。「変換」ボタンはエリア内文字オプションパネル内に配置。グラフィックスタイルの listbox はクリックした時点で選択オブジェクトへ即適用（「適用」ボタンなし）。閉じるボタンは廃止し、パレットをアクティブにして Esc キーで閉じる。worker と重複していた旧・非worker実装を一掃／ Converted the modal dialog into a resident palette (#targetengine + a $.global single-instance guard); DOM work is delegated to the main engine via BridgeTalk and runs while the palette stays open; the "Convert" button now sits inside the Area Type Options panel; clicking a graphic style in the listbox live-applies it to the current selection (no Apply button); the Close button was dropped in favor of pressing Esc while the palette is active; removed the legacy non-worker code duplicated by the workers
- v1.2.0 (2026-07-01): 固定パス（TARGET_FILE_PATH）と固定スタイル名（文字白抜き／枠のみ）を撤去。ダイアログに「スタイルの読み込み」ボタンを追加し、選んだ AI ファイルとスタイル名を Folder.userData に記憶。取り込んだスタイル名からラジオを自動生成する構成へ変更（ImportGraphicStyles v1.7.0 より移植）／ Removed the hardcoded path (TARGET_FILE_PATH) and fixed style names (white text / frame only); added a "Load Styles" button that remembers the picked AI file and style names in Folder.userData; radios are now generated automatically from the imported style names (ported from ImportGraphicStyles v1.7.0)
- v1.1.0 (2026-07-01): ダイアログにグラフィックスタイル選択（元の見た目／文字白抜き／枠のみ）を追加。外部 AI ファイル（TARGET_FILE_PATH）から未登録スタイルを取り込んで適用（ImportGraphicStyles より移植）。角丸オプションを削除／ Added a graphic-style choice (original appearance / white text / frame only); imports the chosen style from an external AI file when it is not yet registered (ported from ImportGraphicStyles). Removed the round-corners option
- v1.0.0 : 初期バージョン / Initial release

### スクリプト情報

- バージョン: v1.3.0
