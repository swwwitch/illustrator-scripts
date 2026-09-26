#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

RGBまたはCMYKで構成された黒を、安定したK100の黒に変換します。
テキスト、パス、スウォッチの塗りおよび線カラーが一括の対象です。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorToK100Converter.md

### Overview

Converts blacks built from RGB or CMYK into a stable K100 black.
Text, paths and swatches are all covered, for both fill and stroke colors.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorToK100Converter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorToK100Converter";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorToK100Converter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorToK100Converter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 置き換え先にするスウォッチ名（無ければ K100 を新規作成） / Swatch used as K100; a new K100 is built when missing */
    var K100_SWATCH_NAME = "ブラック";

    /* 黒と判定する閾値 / Thresholds for treating a color as black */
    var BLACK_THRESHOLDS = {
        rgb: { red: 39, green: 39, blue: 39 },       /* RGB がすべてこれ未満 / All RGB below these */
        cmyk: { cyan: 70, magenta: 70, yellow: 70 }, /* CMY がすべてこれ以上 / All CMY at or above these */
        cmykAll: 60,                                 /* CMYK がすべてこれ以上 / All CMYK at or above this */
        cmykTotal: 310,                              /* CMYK の合計がこれ以上 / CMYK total at or above this */
        solidBlackMaxCmy: 10                         /* 墨ベタとして許容する CMY の上限 / Max CMY kept as solid black */
    };

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
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            rgbDocument: {
                ja: "このスクリプトはCMYKカラーモードのドキュメントでのみ使用できます。",
                en: "This script works only in CMYK color mode documents."
            }
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
    // 色の変換 / Color conversion
    // =========================================

    /**
     * 置き換え先の K100 カラーを返す（スウォッチが CMYK ならその色、無ければ K100 を新規作成）
     * @param {Swatch|null} k100Swatch - 置き換え先のスウォッチ
     * @returns {CMYKColor} K100 のカラー
     */
    function getK100Color(k100Swatch) {
        if (k100Swatch && k100Swatch.color.typename === "CMYKColor") {
            return k100Swatch.color;
        }
        var k100Color = new CMYKColor();
        k100Color.cyan = 0;
        k100Color.magenta = 0;
        k100Color.yellow = 0;
        k100Color.black = 100;
        return k100Color;
    }

    /**
     * RGB の黒か判定する（すべての値が閾値未満）
     * @param {RGBColor} color - 判定するカラー
     * @returns {boolean} 黒なら true
     */
    function isRgbBlack(color) {
        var limits = BLACK_THRESHOLDS.rgb;
        return color.red < limits.red && color.green < limits.green && color.blue < limits.blue;
    }

    /**
     * CMYK のリッチブラックか判定する（墨ベタは除く）
     * @param {CMYKColor} color - 判定するカラー
     * @returns {boolean} K100 に変換する対象なら true
     */
    function isCmykRichBlack(color) {
        var c = color.cyan;
        var m = color.magenta;
        var y = color.yellow;
        var k = color.black;
        var maxCmy = BLACK_THRESHOLDS.solidBlackMaxCmy;

        /* 墨ベタ（K=100 かつ CMY が小さい）はそのまま許容 / Keep solid black as is */
        if (k === 100 && c <= maxCmy && m <= maxCmy && y <= maxCmy) {
            return false;
        }

        /* CMY がすべて閾値以上／CMYK がすべて閾値以上／合計が閾値以上 / CMY, all-CMYK or total over the thresholds */
        var cmyLimits = BLACK_THRESHOLDS.cmyk;
        var allLimit = BLACK_THRESHOLDS.cmykAll;
        return (c >= cmyLimits.cyan && m >= cmyLimits.magenta && y >= cmyLimits.yellow) ||
            (c >= allLimit && m >= allLimit && y >= allLimit && k >= allLimit) ||
            (c + m + y + k) >= BLACK_THRESHOLDS.cmykTotal;
    }

    /**
     * RGB／CMYK の黒なら K100 に置き換えたカラーを返す（それ以外はそのまま）
     * @param {Color} color - 元のカラー
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {Color} 変換後のカラー
     */
    function convertBlackToK100(color, k100Color) {
        if (color.typename === "RGBColor" && isRgbBlack(color)) {
            return k100Color;
        }
        if (color.typename === "CMYKColor" && isCmykRichBlack(color)) {
            return k100Color;
        }
        return color;
    }

    /**
     * パス（複合パス）の塗りと線を変換する。ロック・非表示は対象外
     * @param {PathItem|CompoundPathItem} pathItem - 対象のパス
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {void}
     */
    function convertPathColors(pathItem, k100Color) {
        if (pathItem.locked || pathItem.hidden) return;
        if (pathItem.filled) pathItem.fillColor = convertBlackToK100(pathItem.fillColor, k100Color);
        if (pathItem.stroked) pathItem.strokeColor = convertBlackToK100(pathItem.strokeColor, k100Color);
    }

    /**
     * テキストの文字ごとに塗りと線を変換する。ロック・非表示は対象外
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {void}
     */
    function convertTextColors(textFrame, k100Color) {
        if (textFrame.locked || textFrame.hidden) return;
        var chars = textFrame.characters;
        for (var i = 0; i < chars.length; i++) {
            chars[i].fillColor = convertBlackToK100(chars[i].fillColor, k100Color);
            chars[i].strokeColor = convertBlackToK100(chars[i].strokeColor, k100Color);
        }
    }

    /**
     * グループ内のパス・複合パス・テキストを再帰的に変換する
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {void}
     */
    function convertGroupColors(groupItem, k100Color) {
        var i;
        for (i = 0; i < groupItem.pathItems.length; i++) {
            convertPathColors(groupItem.pathItems[i], k100Color);
        }
        for (i = 0; i < groupItem.compoundPathItems.length; i++) {
            convertPathColors(groupItem.compoundPathItems[i], k100Color);
        }
        for (i = 0; i < groupItem.textFrames.length; i++) {
            convertTextColors(groupItem.textFrames[i], k100Color);
        }
        for (i = 0; i < groupItem.groupItems.length; i++) {
            convertGroupColors(groupItem.groupItems[i], k100Color);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメント内のテキスト・パス・複合パス・グループ・スウォッチの黒を K100 に変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        /* RGB ドキュメントでは中断 / Stop on RGB documents */
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            alert(getLabel("alert.rgbDocument"));
            return;
        }

        /* getByName は見つからないと例外 / getByName throws when the swatch is missing */
        var k100Swatch = null;
        try {
            k100Swatch = doc.swatches.getByName(K100_SWATCH_NAME);
        } catch (e) { }
        var k100Color = getK100Color(k100Swatch);

        var i;
        for (i = 0; i < doc.textFrames.length; i++) {
            convertTextColors(doc.textFrames[i], k100Color);
        }
        for (i = 0; i < doc.pathItems.length; i++) {
            convertPathColors(doc.pathItems[i], k100Color);
        }
        for (i = 0; i < doc.compoundPathItems.length; i++) {
            convertPathColors(doc.compoundPathItems[i], k100Color);
        }
        for (i = 0; i < doc.groupItems.length; i++) {
            convertGroupColors(doc.groupItems[i], k100Color);
        }

        /* スウォッチのカラー定義 / Swatch color definitions */
        for (i = 0; i < doc.swatches.length; i++) {
            var swatchColor = doc.swatches[i].color;
            if (swatchColor.typename === "RGBColor" || swatchColor.typename === "CMYKColor") {
                doc.swatches[i].color = convertBlackToK100(swatchColor, k100Color);
            }
        }
    }

    main();

})();
