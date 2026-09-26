#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの塗りおよび線のカラーをスウォッチ（スポットカラー）として登録し、その場で再適用します。
RGBとCMYKに対応し、既存の同名スウォッチは再利用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterAndApplySwatches.md

### Overview

Registers the fill and stroke colors of the selected objects as spot-color swatches and reapplies them in place.
RGB and CMYK are supported, and an existing swatch with the same name is reused.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegisterAndApplySwatches.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RegisterAndApplySwatches";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterAndApplySwatches.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegisterAndApplySwatches.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },
            noTextFill: { ja: "塗りが設定されていません。", en: "The text has no fill color." },
            applyError: { ja: "カラーの適用中にエラーが発生しました。", en: "An error occurred while applying colors." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文字列を引く
     * @param {string} labelPath - "alert.noSelection" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    // =========================================
    // スウォッチ登録 / Swatch registration
    // =========================================

    /**
     * RGB／CMYK カラーの複製を作る
     * @param {Color} color - 元のカラー
     * @returns {RGBColor|CMYKColor|null} 複製（RGB／CMYK 以外は null）
     */
    function createColorCopy(color) {
        var colorCopy = null;
        switch (color.typename) {
            case "RGBColor":
                colorCopy = new RGBColor();
                colorCopy.red = color.red;
                colorCopy.green = color.green;
                colorCopy.blue = color.blue;
                break;
            case "CMYKColor":
                colorCopy = new CMYKColor();
                colorCopy.cyan = color.cyan;
                colorCopy.magenta = color.magenta;
                colorCopy.yellow = color.yellow;
                colorCopy.black = color.black;
                break;
        }
        return colorCopy;
    }

    /**
     * カラーの値からスウォッチ名を作る（例: "C=0 M=100 Y=100 K=0"）
     * @param {Color} color - 元のカラー
     * @returns {string} スウォッチ名（RGB／CMYK 以外は空文字）
     */
    function buildSwatchName(color) {
        if (color.typename === "RGBColor") {
            return "R=" + Math.round(color.red) + " G=" + Math.round(color.green) + " B=" + Math.round(color.blue);
        }
        if (color.typename === "CMYKColor") {
            return "C=" + Math.round(color.cyan) +
                " M=" + Math.round(color.magenta) +
                " Y=" + Math.round(color.yellow) +
                " K=" + Math.round(color.black);
        }
        return "";
    }

    /**
     * 同名のスポットカラーを探し、無ければ新規に登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} color - 登録するカラー
     * @param {string} swatchName - スウォッチ名
     * @returns {Spot|null} スポットカラー（登録できないカラーは null）
     */
    function findOrAddSpot(doc, color, swatchName) {
        for (var i = 0; i < doc.spots.length; i++) {
            if (doc.spots[i].name === swatchName) {
                return doc.spots[i];
            }
        }
        var colorCopy = createColorCopy(color);
        if (colorCopy === null) {
            return null;
        }
        var spot = doc.spots.add();
        spot.colorType = ColorModel.SPOT;
        spot.color = colorCopy;
        spot.name = swatchName;
        return spot;
    }

    /**
     * 1色をスウォッチ登録し、そのスポットカラーをオブジェクトの塗りまたは線に適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Color} color - 登録するカラー
     * @param {boolean} isStroke - 線に適用するなら true
     * @returns {void}
     */
    function registerAndApplyColor(doc, item, color, isStroke) {
        /* RGB／CMYK 以外（なし・スポット・グレー・グラデーション・パターン）は対象外 / Only RGB and CMYK are registered */
        if (!color) return;
        var swatchName = buildSwatchName(color);
        if (swatchName === "") return;

        var spot = findOrAddSpot(doc, color, swatchName);
        if (!spot) return;

        var spotColor = new SpotColor();
        spotColor.spot = spot;
        spotColor.tint = 100;

        if (isStroke) {
            item.strokeColor = spotColor;
        } else if (item.typename === "TextFrame") {
            item.textRange.characterAttributes.fillColor = spotColor;
        } else {
            item.fillColor = spotColor;
        }
    }

    /**
     * パスの塗りと線、またはテキストの塗りをスウォッチ登録して適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem|TextFrame|CompoundPathItem} targetItem - 対象のオブジェクト
     * @returns {void}
     */
    function registerItemColors(doc, targetItem) {
        if (targetItem.typename === "PathItem") {
            registerAndApplyColor(doc, targetItem, targetItem.fillColor, false);
            registerAndApplyColor(doc, targetItem, targetItem.strokeColor, true);
        } else if (targetItem.typename === "TextFrame") {
            var textFillColor = targetItem.textRange.characterAttributes.fillColor;
            if (!textFillColor || textFillColor.typename === "NoColor") {
                alert(getLabel("alert.noTextFill"));
                return;
            }
            registerAndApplyColor(doc, targetItem, textFillColor, false);
        }
    }

    /**
     * 選択中のオブジェクトを処理する（グループは再帰的にたどる）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {void}
     */
    function processItem(doc, item) {
        if ((item.typename === "PathItem" && item.closed) || item.typename === "TextFrame") {
            registerItemColors(doc, item);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                processItem(doc, item.pageItems[i]);
            }
        } else if (item.typename === "CompoundPathItem") {
            if (item.pathItems.length > 0) {
                registerItemColors(doc, item);
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のオブジェクトのカラーをスウォッチ登録して適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        /* DOM の予期しない失敗を1回の通知にまとめる / Report unexpected DOM failures once */
        try {
            for (var i = 0; i < selectedItems.length; i++) {
                processItem(doc, selectedItems[i]);
            }
        } catch (e) {
            alert(getLabel("alert.applyError") + "\n" + e);
        }
    }

    main();

})();
