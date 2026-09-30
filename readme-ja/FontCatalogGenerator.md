# インストール済みフォントの見本を自動生成

[![Direct](https://img.shields.io/badge/Direct%20Link-FontCatalogGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/FontCatalogGenerator.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontCatalogGenerator.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

システムにインストールされているフォントを一覧化し、アートボード上にフォント見本を自動生成します。

### 主な機能

- 表示文字列とフォントサイズをダイアログで指定
- インストール済みフォントを一覧化して見本を配置

### 使い方

1. スクリプトを実行します。
2. 表示する文字列とフォントサイズを指定します。
3. 実行すると、アートボード上に見本が生成されます。

### 注意点

- フォント数が多い環境では処理に時間がかかります。
- 用途が近いスクリプトとして TypefaceSampler.jsx / FontSampler.jsx があります。

### 更新履歴

- v1.7.8（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.7.7（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.7.6 (2026-09-30) : 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.7.5 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.7.4 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.7.3 (2026-09-29) : 文字の単位が級のとき、単位の表示を「Q/H」から「Q」に修正（単位の扱いを共通の表に一本化）
- v1.7.2 (2026-09-28) : 英語 UI の項目名のコロンの後ろの空白を削除（「:」にそろえた）。ボタン行を共通の部品で組むようにした
- v1.7.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.7.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.6 (2026-01-27)
