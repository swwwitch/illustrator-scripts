#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの塗りおよび線のカラーをスウォッチ（スポットカラー）として登録し、その場で再適用します。
RGBとCMYKに対応し、既存の同名スウォッチは再利用します。

詳細は README を参照してください。

### Overview

Registers the fill and stroke colors of the selected objects as spot-color swatches and reapplies them in place.
RGB and CMYK are supported, and an existing swatch with the same name is reused.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RegisterAndApplySwatches";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-06-26";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterAndApplySwatches.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegisterAndApplySwatches.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

})();
