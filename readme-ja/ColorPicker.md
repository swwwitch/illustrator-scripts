# カラーピッカーの再利用ライブラリ

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/ColorPicker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

他のスクリプトから読み込んで使う、カラーピッカーの再利用ライブラリです。

### 使い方

1. 対象スクリプトの先頭（関数の外）で `#include "ColorPicker.jsx"` と書いて読み込みます。
2. `ColorPicker.show()` を呼び出します。

        var result = ColorPicker.show({
            value: "FF0000",      // "RRGGBB" または "cmyk:C,M,Y,K"
            title: "Color Picker",
            lang: "ja"            // ラベルの言語（"ja" または "en"。省略時は "en"）
        });

3. キャンセルされた場合は `null` が返ります。

### 注意点

- 単体で実行してもダイアログは開きません。
- `#include` を関数の中に書くと、このファイルの `SCRIPT_NAME` / `SCRIPT_VERSION` などがその関数の変数になり、読み込んだ側の同名の値が隠れます。
- 参照元: `jsx/shape/SmartShapeMaker.jsx` / `jsx/stroke-table/LeaderLineBuilder.jsx`
- `jsx/text/AddBulletsAndNumbers.jsx` は v1.2.2 から本体に取り込んでおり、このファイルを読み込みません。

### 更新履歴

- v1.1.5（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.1.2 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.1.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0.2 (2026-09-21) : スライダーにツールチップの言語が渡っておらず、ピッカーを開くとエラーになる不具合を修正
- v1.0.1 (2026-09-19) : 各コントロールにツールチップを追加
- v1.0
