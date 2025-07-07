#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

RegisterAndApplySwatches.jsx

### 概要

- 選択オブジェクト（閉じたパス、テキスト）の塗りおよび線のカラーをスウォッチ（スポットカラー）として登録し、その場で再適用するIllustrator用スクリプトです。
- RGBおよびCMYKカラーをサポートし、既存の同名スウォッチを再利用します。

### 主な機能

- RGB または CMYK カラーのスウォッチ登録
- 同名スウォッチがある場合は再利用
- 塗り・線にスポットカラーを即時適用
- テキストオブジェクトの文字カラーにも対応
- グループや複合パス内のオブジェクトを再帰処理

### 処理の流れ

1. 対象オブジェクトを選択
2. RGB または CMYK カラーを検出
3. スウォッチ名を生成し、既存スウォッチがあれば再利用
4. スポットカラーを作成してオブジェクトに適用

### 更新履歴

- v1.0.0 (20250626) : 初期バージョン

---

### Script Name:

RegisterAndApplySwatches.jsx

### Overview

- An Illustrator script to register fill and stroke colors of selected objects (closed paths, text) as swatches (spot colors) and reapply them immediately.
- Supports RGB and CMYK colors, and reuses existing swatches with the same name.

### Main Features

- Register swatches from RGB or CMYK colors
- Reuse existing swatches with matching names
- Immediately apply spot colors to fills and strokes
- Supports text object character colors
- Recursively process groups and compound paths

### Process Flow

1. Select target objects
2. Detect RGB or CMYK colors
3. Generate swatch names and reuse existing swatches if found
4. Create and apply spot colors to objects

### Update History

- v1.0.0 (20250626): Initial version
*/

function createColorCopy(color) {
    var c = null;
    switch (color.typename) {
        case "RGBColor":
            c = new RGBColor();
            c.red = color.red;
            c.green = color.green;
            c.blue = color.blue;
            break;
        case "CMYKColor":
            c = new CMYKColor();
            c.cyan = color.cyan;
            c.magenta = color.magenta;
            c.yellow = color.yellow;
            c.black = color.black;
            break;
    }
    return c;
}

function registerColorToSwatch(targetItem) {
    function registerSingleColorToSwatch(item, color, isStroke) {
        if (!color || color.typename === "NoColor" || color.typename === "SpotColor" || color.typename === "GrayColor" || color.typename === "GradientColor" || color.typename === "PatternColor") {
            return false; // スルー
        }

        var swatchName = "";
        if (color.typename === "RGBColor") {
            swatchName = "R=" + Math.round(color.red) + " G=" + Math.round(color.green) + " B=" + Math.round(color.blue);
        } else if (color.typename === "CMYKColor") {
            swatchName =
                "C=" + Math.round(color.cyan) +
                " M=" + Math.round(color.magenta) +
                " Y=" + Math.round(color.yellow) +
                " K=" + Math.round(color.black);
        } else {
            // alert("対応していないカラーモデルです。RGBまたはCMYKのみ対応しています。");
            return false;
        }

        var spot = null;
        for (var i = 0; i < app.activeDocument.spots.length; i++) {
            if (app.activeDocument.spots[i].name === swatchName) {
                spot = app.activeDocument.spots[i];
                break;
            }
        }

        if (!spot) {
            var dupColor = createColorCopy(color);
            if (dupColor === null) {
                // alert("対応していないカラーモデルです。RGBまたはCMYKのみ対応しています。");
                return false;
            }

            spot = app.activeDocument.spots.add();
            spot.colorType = ColorModel.SPOT;
            spot.color = dupColor;
            spot.name = swatchName;
        }

        var spotColor = new SpotColor();
        spotColor.spot = spot;
        spotColor.tint = 100;

        if (isStroke) {
            item.strokeColor = spotColor;
        } else {
            if (item.typename === "TextFrame") {
                item.textRange.characterAttributes.fillColor = spotColor;
            } else {
                item.fillColor = spotColor;
            }
        }

        return true;
    }

    if (targetItem.typename === "PathItem") {
        var fillColor = targetItem.fillColor;
        var strokeColor = targetItem.strokeColor;

        var fillResult = false;
        var strokeResult = false;

        if (fillColor && fillColor.typename !== "NoColor") {
            fillResult = registerSingleColorToSwatch(targetItem, fillColor, false);
        }

        if (strokeColor && strokeColor.typename !== "NoColor") {
            strokeResult = registerSingleColorToSwatch(targetItem, strokeColor, true);
        }

        if (!fillResult && !strokeResult) {
            return false; // 警告なしでスキップ
        }

        return fillResult || strokeResult;

    } else if (targetItem.typename === "TextFrame") {
        var fillColor = targetItem.textRange.characterAttributes.fillColor;
        if (!fillColor || fillColor.typename === "NoColor") {
            alert("塗りが設定されていません。");
            return false;
        }
        return registerSingleColorToSwatch(targetItem, fillColor, false);
    } else {
        return false;
    }
}

function processItem(item) {
    if ((item.typename === "PathItem" && item.closed) || item.typename === "TextFrame") {
        registerColorToSwatch(item);
    } else if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            processItem(item.pageItems[i]);
        }
    } else if (item.typename === "CompoundPathItem") {
        if (item.pathItems.length > 0) {
            registerColorToSwatch(item);
        }
    }
}

function main() {
    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
        alert("オブジェクトを選択してください。");
        return;
    }

    var sel = app.activeDocument.selection;

    try {
        for (var i = 0; i < sel.length; i++) {
            processItem(sel[i]);
        }
    } catch (e) {
        alert("カラーの適用中にエラーが発生しました:\n" + e);
    }
}

main();