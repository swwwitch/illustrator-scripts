# ドキュメントの切替

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSwitchDocs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/SmartSwitchDocs.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSwitchDocs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

複数のドキュメントが開いているときに、別のドキュメントへすばやく切り替えます。

<img alt="" src="https://github.com/user-attachments/assets/e2e98c44-0db3-46c6-9f0e-579b17b82599" width="70%" />

### 主な機能

- 開いているドキュメントが**1つ**なら、何もしない
- **2つ**なら、ダイアログを表示せず、アクティブでない方へ切り替える
- **3つ以上**なら、ダイアログのリストから切り替え先を選ぶ（起動時のドキュメントはリストから除き、［元のドキュメント］に表示）
- ［プレビュー］（既定オン）がオンなら、リストで選んだ時点で切り替える。オフなら［OK］をクリックしてから切り替える
- リストの項目をダブルクリックすると、［OK］と同じく切り替えて閉じる
- ［キャンセル］・Esc・閉じるボタンで閉じると、元のドキュメントに戻る
- 実行時にドキュメントウィンドウをタブにまとめる（［すべてのウィンドウを統合］）

### 使い方

1. ドキュメントを2つ以上開いた状態でスクリプトを実行する
2. 3つ以上のときは、↑↓キーまたはクリックで切り替え先を選ぶ
3. ［OK］またはダブルクリックで確定する。［キャンセル］なら元のドキュメントに戻る

### 紹介記事

- [【Illustrator】ドキュメント切替をスムーズに行うスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/nd9c7b7c077fb)

### 謝辞

- 矢印キーでの連続切り替え、表示時の即時切り替えなどは、三枝 優介さんの改良を取り入れています
  - [uske-s.hatenablog.com](https://uske-s.hatenablog.com/entry/2025/04/03/113115)

### 更新履歴

- v1.0.0 (20250325) : 初期バージョン
- v1.1.0 (20250403) : 矢印キーで選択したあともダイアログにフォーカスを戻して連続で切り替えられるよう修正、表示時に最初の候補へ切り替え、レイアウトを preferredSize に変更（三枝 優介さんによる改良）
- v0.5.1 (20250525) : キャンセルボタンを追加、UIを調整
- v0.5.2 (20250525) : 矢印キーで選択したあとのフォーカス維持を修正
- v0.5.3 (20260903) : ［プレビュー］チェックボックスを追加（オフのときは［OK］をクリックしてから切り替え）、ダイアログタイトルにバージョンを表記、ドキュメントが2つのときにアクティブでない方へ確実に切り替わるよう修正。あわせて基本情報に紹介記事URLを追加、LABELSをカテゴリ入れ子＋`getLabel()`に整理、変数・パネル・関数名を命名規約に統一、ダイアログ構築とドキュメント収集を関数分割、選択中ドキュメントの取得・切り替えの重複を統合、全関数にJSDocを付与
- v0.5.4 (20260924) : ウィンドウの閉じるボタンで閉じたときも元のドキュメントへ戻るよう修正、パネル名「現在のドキュメント」を「元のドキュメント」に変更してツールチップを追加、ダイアログ構築の一部を関数に分割
- v0.5.5 (20260924) : リストの項目をダブルクリックすると切り替えて閉じるよう対応、リストにツールチップを追加
