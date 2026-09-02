# ドキュメントの切替

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartSwitchDocs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/SmartSwitchDocs.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

- 開いているドキュメント数が**1つ**なら、何も起きない
- 開いているドキュメント数が**2つ**なら、ダイアログボックスを表示せず、アクティブでないドキュメントに切り換える
- 開いているドキュメント数が**3つ**以上なら、ダイアログボックスでドキュメントを切り換える（ただし、現在のアクティブなドキュメントは別扱いにする）
- ［プレビュー］（既定オン）をオフにすると、リストで選んだ時点では切り替えず、［OK］をクリックしてから切り替える

<img alt="" src="https://github.com/user-attachments/assets/e2e98c44-0db3-46c6-9f0e-579b17b82599" width="70%" />

### 補足記事：

- [【Illustrator】ドキュメント切替をスムーズに行うスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/nd9c7b7c077fb)

### 更新履歴

- v1.0.0 (20250325) : 初期バージョン
- v0.5.1 (20250525) : キャンセルボタンを追加、UIを調整
- v0.5.2 (20250525) : 矢印キーで選択したあとのフォーカス維持を修正
- v0.5.3 (20260903) : ［プレビュー］チェックボックスを追加（オフのときは［OK］をクリックしてから切り替え）、ダイアログタイトルにバージョンを表記、ドキュメントが2つのときにアクティブでない方へ確実に切り替わるよう修正。あわせて基本情報に紹介記事URLを追加、LABELSをカテゴリ入れ子＋`getLabel()`に整理、変数・パネル・関数名を命名規約に統一、ダイアログ構築とドキュメント収集を関数分割、選択中ドキュメントの取得・切り替えの重複を統合、全関数にJSDocを付与
