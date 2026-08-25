# register-temp-style

[![Direct](https://img.shields.io/badge/Direct%20Link-register--temp--style.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/register-temp-style.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/register-temp-style.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択中のオブジェクトの見た目を「名称未設定」で登録する Illustrator 標準の挙動を回避し、固定名のグラフィックスタイルとして登録するスクリプト。
- 同名スタイルが既に存在する場合は削除してから上書き登録するので、繰り返し実行しても重複しない。
- Registers the appearance of the selected object as a graphic style with a fixed name, bypassing Illustrator's default "Untitled" registration.
- If a style with the same name already exists, it is removed first so repeated runs do not create duplicates.

### 処理の流れ

1. 選択が1つでなければ何もせず終了 / Exit silently unless exactly one object is selected.
2. 既存の同名スタイルを削除 / Remove the existing style with the same name, if any.
3. 一時アクションをロード→実行→アンロードして、無名のグラフィックスタイルを末尾に追加 / Load, run, and unload a temporary action that appends an unnamed graphic style.
4. 末尾のスタイルを `TEMP_STYLE_NAME` に改名 / Rename the last style to `TEMP_STYLE_NAME`.

### 設定 / Settings

- `TEMP_STYLE_NAME` を書き換えると登録名を変更できる / Edit `TEMP_STYLE_NAME` to change the registered style name.

### オリジナルアイデア

@comsk(asa me)さん
https://qiita.com/comsk/items/87161b2b7d2336b161c4

### スクリプト情報

- バージョン: v1.0.0
