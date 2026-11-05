# パネルを検索して表示

[![Direct](https://img.shields.io/badge/Direct%20Link-AiPanelLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiPanelLauncher.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiPanelLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

［ウィンドウ］メニューのパネルを、サブメニュー（［書式］）の中まで展開して1つのリストに並べます。検索欄で絞り込み、選んだパネルを表示します。InDesign 版 [IdPanelLauncher](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdPanelLauncher.md) の Illustrator 版です。

### 主な機能

- ［ウィンドウ］メニューのパネルを「名前｜分類」の2列のリストで表示
  - ［書式］の中の項目（文字、段落、OpenType など）も展開し、分類の列に「書式」と表示
  - 使っている Illustrator にまだ無いパネル（Retype、コンセプトからベクター生成など）は出さない
- 検索欄で名前と分類を絞り込み（入力するたびに絞り込む）
  - 空白で区切ると、すべての語を含むものに絞る
  - 大文字小文字を区別しない。ひらがなはカタカナとして探し、漢字の名前は読みでも探す（「しょしき」で［書式］の中のパネル）
  - 検索欄の右の（×）ボタンで検索をクリア
- セットのボタンで、決まった組み合わせのパネルをまとめて表示（検索・選択に関係なく、ダイアログを閉じて表示）
  - **テキスト**：文字・段落・文字スタイル・段落スタイル・字形
  - **アピアランス**：アピアランス・グラフィックスタイル
  - **カラー**：カラー・スウォッチ・カラーガイド・グラデーション
- 並び順をラジオボタンで切り替え
  - **デフォルト**：［ウィンドウ］メニューの並び（［書式］の中身は［書式］の位置）
  - **あいうえお順**：パネル名の読みの順（英字が先）
  - 並び替えても、選んでいたパネルは選んだまま
- 絞り込むと先頭のパネルを選択するので、そのまま Enter で表示できる
- 複数のパネルを選んでまとめて表示
- ダイアログのタイトルに表示中の件数を表示。ダイアログの位置は次回も同じ場所に開く

### 使い方

1. スクリプトを実行する
2. 検索欄にパネルの名前の一部を入れる
3. 次のどれかで表示する
   - Enter／Return
   - リストの項目をダブルクリック
   - ［表示］をクリック

検索欄で ↓ を押すとリストへ移ります。⌘（Ctrl）／Shift＋クリックで複数選べます。

### 制限事項・メモ

- Illustrator のスクリプトからはメニューを読めないため、パネルの一覧（名前・メニューコマンド ID・対応バージョン）はスクリプト内の表（PANEL_TABLE）に持っています。Illustrator 2026（30.8）日本語版の［ウィンドウ］メニューに合わせています。エクステンションやプラグインのパネル、各種ライブラリは出しません。
- ［ブレンド］パネル（30.7〜）はメニューコマンド ID が分からないため、一覧に入れていません。
- 表示は［ウィンドウ］メニューの項目を選ぶのと同じです（`app.executeMenuCommand()`）。Illustrator のメニューは表示／非表示の切り替えで、スクリプトからは今の状態を読めないため、開いているパネルを選ぶと閉じることがあります。
- パネルはダイアログを閉じてから表示します。
- セットのボタンも1つずつメニューの項目を選ぶのと同じなので、セットの中にすでに開いているパネルがあると、そのパネルは閉じることがあります。組み合わせはスクリプト内の PANEL_SETS で変えられます。
- 英語の UI では、［ウィンドウ］メニューがアルファベット順なので、デフォルトもあいうえお順（名前順）と同じ並びになります。
- あいうえお順の漢字の読みは、スクリプト内の表（KANJI_READINGS）に持っています。長音（ー）と中黒（・）は無視します。
- 繰り返し使う前提で `#targetengine` を指定しています。キーボードショートカットを割り当てると便利です。
- Esc でダイアログを閉じます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/misc/AiPanelLauncher.jsx` |
| バージョン | v1.1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-05 |
| 最終更新 | 2026-11-06 |
| 動作確認 | Illustrator 2026（30.8）／ macOS |

### 更新履歴

- v1.1.0（2026-11-06）：決まった組み合わせのパネルをまとめて表示するセットのボタン（テキスト・アピアランス・カラー）を追加
- v1.0.0（2026-11-05）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
