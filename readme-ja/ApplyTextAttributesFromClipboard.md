# 保存した文字属性をテキストに適用

[![Direct](https://img.shields.io/badge/Direct%20Link-ApplyTextAttributesFromClipboard.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ApplyTextAttributesFromClipboard.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyTextAttributesFromClipboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

CopyTextAttributesToClipboard.jsx が保存した文字属性を、選択中のテキストへ適用します。

永続エンジン "FontClipboard" の `$.global.FontClipboard` から読み取るため、2つのスクリプトを組み合わせて「書式のコピー＆ペースト」として使います。

### 主な機能

- フォント・サイズ／行送り、カーニング関連、段落属性、塗りとグラフィックスタイルの4パネル構成
- 各属性はチェックボックスで適用可否を切り替え（初期状態でONなのはフォントのみ）
- 「塗りとグラフィックスタイル」パネルはラジオで排他（しない／塗り／グラフィックスタイル）

### 使い方

1. CopyTextAttributesToClipboard.jsx で属性をコピーしておきます。
2. 適用先のテキストを選択します（文字ツールで部分選択しても構いません）。
3. スクリプトを実行し、適用する属性にチェックを入れて［OK］をクリックします。

### 注意点

- テキスト編集モードで部分選択している場合は、その範囲だけに適用します。
- 複数の TextFrame を選択している場合は、それぞれに同じ属性を適用します。
- 属性の受け渡しには `#targetengine "FontClipboard"` を使うため、Illustratorを終了すると内容は失われます。

### 更新履歴

- v1.3.1 (2026-05-21)
- v1.3.3 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.4 (2026-09-28) : 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした。ボタン行を共通の部品で組むようにした
- v1.3.5 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.3.6 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.3.7（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.3.8（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.3.9（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
