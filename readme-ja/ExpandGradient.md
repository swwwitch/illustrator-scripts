# ExpandGradient

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandGradient.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandGradient.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandGradient.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトのグラデーションを、指定した数の単色オブジェクトに分割します。

分割後の後処理として、重なりを整理してひとまとめにする、両端からブレンドを作成する、のいずれかを選べます。

### 使い方

1. グラデーションを適用したオブジェクトを選択します。
2. スクリプトを実行します。
3. ステップ数と実行後の処理を指定して［OK］をクリックします。

### オプション

**ステップ数**

分割後にできる単色オブジェクトの数。既定値 5、2 以上の整数。↑↓キーで±1、Shift+↑↓で10の倍数にスナップします。

「ブレンドに変換」を選んだときは 2 に固定され、入力欄はディム表示になります。

**実行後の処理**

- なし: ［オブジェクト］＞［分割・拡張］のみ（クリッピンググループのまま）
- 単純に拡張: パスファインダーのクロップとマージを適用し、単色オブジェクトの集合に整理
- ブレンドに変換: 上記のあと、両端のオブジェクトだけを残してブレンドを作成

### 注意点

- 内部で一時的な .aia アクションを生成・ロード・再生し、実行後にアンロードして一時ファイルを削除します。アクションのセット名／アクション名は定数で管理し、.aia 内の名前も同じ値から生成します。
- Illustrator の分割・拡張は指定値より1つ少ない数を返すため、スクリプト側で補正しています。
- 「ブレンドに変換」のステップ数は、Illustrator のブレンドオプションの現在値がそのまま使われます。

### 紹介記事

[【Illustrator】グラデーションを単色のカラーのオブジェクトに分割する｜DTP Transit 別館](https://note.com/dtp_tranist/n/nbe084e691ba5)

### 更新履歴

- v1.1.2 (2026-08-27): 分割数が指定より1つ少なくなる問題を修正、「ブレンドに変換」時はステップ数を2に固定してディム表示、紹介記事へのリンクを追加、ExpandGradient-v2.jsx を削除、コードを整理
- v1.1.0 (2026-05-25)
