# ColorPicker

[![Direct](https://img.shields.io/badge/Direct%20Link-ColorPicker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/stroke-table/ColorPicker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorPicker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

他のスクリプトから読み込んで使う、カラーピッカーの再利用ライブラリです。

### 使い方

1. 対象スクリプトから `#include "ColorPicker.jsx"` で読み込みます。
2. `ColorPicker.show()` を呼び出します。

        var result = ColorPicker.show({
            value: "FF0000",      // "RRGGBB" または "cmyk:C,M,Y,K"
            title: "Color Picker"
        });

3. キャンセルされた場合は `null` が返ります。

### 注意点

- 単体で実行してもダイアログは開きません。
- 使用例: `jsx/shape/SmartShapeMaker.jsx`
- `jsx/text/ColorPicker.jsx` は同一内容のファイルです。

### 更新履歴

- v1.0
