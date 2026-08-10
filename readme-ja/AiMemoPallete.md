# AiMemoPallete

[![Direct](https://img.shields.io/badge/Direct%20Link-AiMemoPallete.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AiMemoPallete.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiMemoPallete.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

Illustrator 用のメモ入力フローティングパレット。

### 読み込み

- 選択オブジェクトからテキストを収集（**追加**＝既定 / **置き換え**）。追加時は既存テキストとの間に空行を1つ挟む
- グループ・クリップグループ・シンボル内（入れ子も多段展開）のテキストも対象
- 選択が無いときはクリップボードを一時的に貼り付け（`pasteInAllArtboard`）→テキスト収集→貼り付けたオブジェクトを削除、の手順で読み込む（アートボードごとの複製は重複除去）
- 複数取得時はカンバス上の位置で上から順（同じ高さは左から右）に整列。実質的に空のフレームは無視
- 取り込み時に濁点・半濁点を NFC へ補正
- 各ボタンにはツールチップ（helpTip）を表示

### 編集・コピー

- 「空行削除」でテキスト欄の空行（空白のみの行を含む）を除去
- 「改行削除」でテキスト欄の改行をすべて除去し1行にまとめる
- 「すべてをコピー」でテキスト欄の内容をシステムクリップボードへコピー（一時テキストフレーム + `app.copy()` を BridgeTalk 実行）
- テキストが空／対象が無いときは該当ボタンを自動でディム（空行が無ければ「空行削除」、改行が無ければ「改行削除」もディム）

### 保存・その他

- 入力したメモを UTF-8 のテキストファイル（.txt）へ保存（保存元・保存日時のフッター付き）
- 保存先は `SAVE_LOCATION_MODE` で切替：保存ダイアログで選択（A）/ 常にデスクトップへ `memo-<ドキュメント名>-<yyyymmdd>.txt`（B＝既定）
- 保存後はパレットを閉じず、入力内容とウィンドウ位置をそのまま保持（再起動しない）
- クリア後はスクリプトを再起動して空の状態で起動し直す
- ボタンエリアはテキスト欄の上＝読み込み系（読み込む／空行削除／改行削除）、下＝3カラム（左:保存・すべてをコピー／中央:スペーサー／右:クリア）
- 閉じたときの挙動は `CLEAR_ON_CLOSE` で切替：内容を保持（既定）/ クリア。ウィンドウ位置は常に保持
- UI・メッセージ・保存フッターまで日本語／英語にローカライズ（ロケール自動判定）

### 補足

選択テキストの取得は、生きた DOM を持つメインエンジンへ BridgeTalk で問い合わせて行う
（常駐エンジンの app はパレット表示中に DOM 接続を失い `there is no document` を投げるため）。
シンボル内テキストは一時レイヤーへ複製→breakLink で展開して読み取り、複製は削除して元のシンボルには触れない。
Illustrator には `app.system` が無いため、クリップボードへのコピーは一時テキストフレーム + `app.copy()` で行う。
逆方向（クリップボードの取り込み）も `app.system` で直接読めないため、ドキュメントへ一時的に貼り付けてからテキストを読み取り、貼り付けたオブジェクトは削除する（元の選択は復元）。

### 解説

https://note.com/dtp_tranist/n/n41e91e4b1a09

### オリジナル

こじらせたクマーさんの以下の記事をもとに、機能追加やリファクタリングを行いました。
https://note.com/nice_lotus120/n/n6291a432b30d

### スクリプト情報

- バージョン: v1.1.2
