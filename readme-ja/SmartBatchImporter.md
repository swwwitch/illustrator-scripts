# 開いているファイルを1つに整列統合

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartBatchImporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/files/SmartBatchImporter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBatchImporter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

- 複数の AI/SVG ファイルをバッチ処理で読み込み、表示中かつロック解除されているオブジェクトを対象に
- ファイルが開いているときには、そのファイルが対象。そうでない場合には、読み込みフォルダーを指定
- 新規ドキュメントに貼り付け、グループ化してアートボード内に自動整列します。
- 各グループにはオプションでファイル名ラベルを追加可能（拡張子の表示有無も指定）
- 読み込み後のドキュメントを閉じる／閉じないを指定可能

[【Illustrator】開いているファイルを1つに整列統合するIllustratorスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n8180588e5630)

### 更新履歴

- v1.4.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.4.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.4.2 (20260928) : 英語表示で、ペースト失敗メッセージのコロンを全角から半角（後ろに空白）にした
- v1.4.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.4.2 (20260928) : クリップグループの範囲をマスクで測るようにした（アートボード内の位置合わせとセルの大きさに効く）
- v1.4.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.4.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
