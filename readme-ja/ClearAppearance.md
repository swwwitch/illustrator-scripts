# アピアランスを消去して塗り・線だけ残す

[![Direct](https://img.shields.io/badge/Direct%20Link-ClearAppearance.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ClearAppearance.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearAppearance.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトに「アピアランスの消去」を実行します。

実行前にダイアログを表示し、復元するかどうかと、選択オブジェクトに応じた復元方法を選べます。

### 主な機能

- パス・複合パスは元の塗り・線・線幅を再適用
- 復元方法を「塗り・線・線幅」「塗り・線＋線の詳細設定」から選択
- 「塗り・線＋線の詳細設定」では線端・角の形状・破線・破線オフセット・角の比率も復元
- 「不透明度」「描画モード」「オーバープリント」を個別にON/OFF

### 使い方

1. アピアランスを消去したいオブジェクトを選択します。
2. スクリプトを実行します。
3. 復元方法を指定して［OK］をクリックします。

### 紹介記事

https://note.com/dtp_tranist/n/na4c70c5acd60

### 更新履歴

- v1.0.9（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.0.8（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.0.7（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.0.6 (2026-09-30): 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.0.5 (2026-09-29): ダイアログの不透明度を98%に変更
- v1.0.4 (2026-09-28): 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした。日本語表示で、失敗件数と詳細行のコロンを全角にそろえた。ボタン行を共通の部品で組むようにした
- v1.0.3 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.0.1 (2026-09-18)
- v1.0 (2026-04-14)
