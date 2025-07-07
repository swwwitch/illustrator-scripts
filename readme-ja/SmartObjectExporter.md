# 選択したオブジェクトを書き出すスクリプト

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartObjectExporter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/SmartObjectExporter.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


選択したオブジェクトを一時アートボードに収めて PNG 書き出しするスクリプトです。
書き出し倍率・余白・背景・罫線・ファイル名・保存先などを UI で柔軟に設定でき、プレビュー機能によりリアルタイムで結果を確認できます。

- 選択オブジェクトを囲む一時アートボードを生成し、そこから PNG 書き出し
- 背景は透過／白／黒／任意カラー、透明グリッドから選択可能
- 書き出し倍率（1x/2x/3x/4x/カスタム%）または幅(px)指定が可能
- 余白と罫線の有無・サイズ・色を設定でき、単位に応じて自動換算
- ファイル名は接尾辞＋記号（-／_）の組み合わせで動的に構成
- 禁則文字（¥ / : * ? " < > | スペース類）は選択記号に自動置換
- 書き出し先はデスクトップ／元ファイルと同じ場所を選択可能
- 書き出し後、自動的に Finder/Explorer で保存フォルダを開く
- プリセットの保存・読み込み機能あり（倍率、背景、余白、接尾辞などを記憶）
- ロックされたオブジェクトや非表示レイヤーも適切に処理
- 書き出し対象外のオブジェクトは一時的に非表示にして安全に処理
- 書き出し位置・サイズはピクセルパーフェクトに調整（XYWH を整数化）
- ファイル名を参照する／しないを設定できるようにしました。

[【Illustrator】選択したオブジェクトを書き出すスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/necf308c39f5d)
